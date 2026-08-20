#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
csvtools.py —— 一批大型 CSV 交付檔的三件事: 去重、轉 parquet、看懂附的腳本。

放在那批資料的主資料夾底下跑。三個模式各自獨立:

    python csvtools.py --scripts
        把 misc/ 裡副檔名是 .txt 但其實是 Python 的檔解析出來:
        它讀什麼、寫什麼、路徑寫死在哪一行、怎麼跑。秒級。

    python csvtools.py --dup
        兩個資料夾的檔是不是同一批。先比大小 (秒級) 再比 SHA256 (要讀檔)。
        --quick 只比大小。

    python csvtools.py --parquet --dry
        估算轉成 parquet 能省多少空間 (每個檔只讀前面幾萬列來推估)。
    python csvtools.py --parquet
        真的轉。原檔不動, 寫到 parquet/ 底下。

需要 pyarrow 才能轉 parquet (pip install pyarrow)。其餘模式不需要。
"""
from __future__ import annotations

import argparse
import ast
import csv
import datetime as dt
import hashlib
import os
import re
import sys
import time
from collections import Counter
from pathlib import Path

# ============================================================ 設定 (寫死)
# 改一次就往上加一號。畫面第一行會印出來, 加上腳本自己的絕對路徑 —— 手邊
# 同時存在新舊兩份的時候, 光看輸出根本分不出跑的是哪一個, 之前已經因為這
# 個白跑過一輪。
BUILD = "2026-09-18 #59  (--cmp --loose 認得逗號分隔與帶空格的多值欄)"
MISC = "misc"                 # 放對方的產出與雜項
OUT_PARQUET = "parquet"       # 轉檔輸出到這裡
SAMPLE_ROWS = 50_000          # --dry 每個檔取樣幾列來推估壓縮比
REPORT = "csvtools_report.txt"
# ============================================================ 設定結束

try:
    from tqdm import tqdm
except ImportError:      # 沒裝就照跑, 只是沒有進度條
    class _Null:
        def update(self, *a):
            pass

        def close(self):
            pass

    def tqdm(x=None, **k):
        return x if x is not None else _Null()

MON = "jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec"
_lines: list[str] = []


def add(s: str = "") -> None:
    _lines.append(s)
    print(s, flush=True)


def mb(n: int) -> str:
    return f"{n / 1e6:,.1f}MB" if n < 1e9 else f"{n / 1e9:,.2f}GB"


def series_of(name: str) -> str:
    """檔名去掉數字與月份之後剩下的字 = 系列名。

    順序要緊: 先去數字。"1-15aug2017" 這種月份與數字黏在一起的寫法,
    \\b 邊界在 "5" 與 "a" 之間不成立, 先拆月份會拆不掉, 同一個系列就被
    按月拆成十幾個。
    """
    n = name.lower()
    for ext in (".csv", ".txt", ".xlsx", ".parquet"):
        if n.endswith(ext):
            n = n[: -len(ext)]
    n = re.sub(r"[\d_\-.()]+", " ", n)
    n = re.sub(rf"\b({MON})[a-z]*\b", " ", n)
    n = re.sub(r"\b(1h|2h|1q|2q|3q|4q|all|onward)\b", " ", n)
    n = re.sub(r"\b[hq]\b", " ", n)   # "1H2017" 去掉數字之後剩下的孤字
    return " ".join(n.split()) or "(無)"


def parse_span(name: str) -> tuple[set, str] | tuple[None, None]:
    """檔名 -> (涵蓋哪些 (年,月), 顆粒度標籤)。認不出回 (None, None)。

    這批檔的顆粒度很雜: 半月 / 整月 / 月份範圍 / 季 / 半年 / 整年。要判斷
    「跑的是不是完整的」就得先把每個檔攤成它涵蓋的月份, 再看有沒有洞。
    """
    n = name.lower()
    for ext in (".csv", ".parquet", ".xlsx", ".txt"):
        if n.endswith(ext):
            n = n[: -len(ext)]
    # 底線當分隔字 —— "202304_flown_pax_tax_01" 裡的底線是 \w, \b 邊界不成立,
    # 不換掉的話底下每一條規則都認不出這個檔。
    n = n.replace("(1)", " ").replace("_", " ")

    def months(y, a, b):
        return {(y, m) for m in range(a, b + 1)}

    # 半年 / 季 / 整年
    m = re.search(r"\b([12])h\s*(20\d{2})|\b(20\d{2})\s*([12])h\b", n)
    if m:
        h = int(m.group(1) or m.group(4))
        y = int(m.group(2) or m.group(3))
        return months(y, 1 if h == 1 else 7, 6 if h == 1 else 12), f"{h}H"
    m = re.search(r"\b([1-4])q\s*(20\d{2})|\b(20\d{2})\s*([1-4])q\b", n)
    if m:
        q = int(m.group(1) or m.group(4))
        y = int(m.group(2) or m.group(3))
        return months(y, q * 3 - 2, q * 3), f"{q}Q"

    # 月份範圍  Sep-Oct2014 / Jul-Aug2015
    m = re.search(rf"({MON})[a-z]*\s*-\s*({MON})[a-z]*\s*(20\d{{2}})", n)
    if m:
        a, b = MONTHS[m.group(1)], MONTHS[m.group(2)]
        return months(int(m.group(3)), a, b), "月範圍"

    # 半月  1-15Aug2017 / 16-31Aug2017 / 13-30Nov2014
    m = re.search(rf"(\d{{1,2}})\s*-\s*(\d{{1,2}})\s*({MON})[a-z]*\s*(20\d{{2}})", n)
    if m:
        return ({(int(m.group(4)), MONTHS[m.group(3)])},
                "上半月" if int(m.group(1)) <= 1 else
                ("下半月" if int(m.group(1)) >= 16 else "部分月"))

    # ita 的半月寫法: 年月黏在一起再接 1/2  "2017Aug1" = 八月上半。要擺在
    # 單月規則前面, 否則 "2017Aug" 先被當成整月, 尾巴那個 1 就沒人看了。
    m = re.search(rf"(20\d{{2}})({MON})[a-z]*([12])(?!\d)", n)
    if m:
        return ({(int(m.group(1)), MONTHS[m.group(2)])},
                "上半月" if m.group(3) == "1" else "下半月")

    # 單月  Aug2017 / 2019 Apr / 2019 JUL
    m = re.search(rf"(20\d{{2}})\s*({MON})[a-z]*|({MON})[a-z]*\s*(20\d{{2}})", n)
    if m:
        y = int(m.group(1) or m.group(4))
        mo = MONTHS[m.group(2) or m.group(3)]
        return {(y, mo)}, "整月"

    # YYYYMM  xxxx 202209 xxxx… / 202304 …01
    m = re.search(r"(?:^|\D)(20\d{2})(0[1-9]|1[0-2])(?:\D|$)", n)
    if m:
        return {(int(m.group(1)), int(m.group(2)))}, "整月"

    # 只有年份
    m = re.search(r"\b(20\d{2})\b", n)
    if m:
        y = int(m.group(1))
        return months(y, 1, 12), "整年"

    # 月份 + 兩位年  SEP14 / SEP20。擺在最後 —— 兩位數字也可能是日期, 所以
    # 只有前面所有規則(含四位年份)都認不出來才用, 而且要落在合理年份區間。
    m = re.search(rf"\b({MON})[a-z]*\s*([0-2]\d)\b", n)
    if m:
        y = 2000 + int(m.group(2))
        if 2010 <= y <= 2029:
            return {(y, MONTHS[m.group(1)])}, "整月"
    return None, None


MONTHS = {m: i for i, m in enumerate(
    ["jan", "feb", "mar", "apr", "may", "jun",
     "jul", "aug", "sep", "oct", "nov", "dec"], 1)}


def csv_header(p: Path) -> list[str]:
    """表頭原樣取出 —— **不要 strip**。

    pyarrow 用的是原樣的欄名 (可能帶前後空白)。這邊 strip 掉再拿去當
    column_types 的鍵就對不上, 那一欄會退回型別推斷, 票號 "0000000000000"
    直接變成 0 —— 而且不會報錯, 靜靜地失真。用 csv.reader 而不是 split(",")
    也是同一個道理: 欄名裡可能有引號包住的逗號。
    """
    with open(p, "rb") as f:
        line = f.readline().rstrip(b"\r\n").decode("utf-8", "replace")
    return next(csv.reader([line.lstrip("\ufeff")]))


def sha256(p: Path, bar=None) -> str:
    h = hashlib.sha256()
    with open(p, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
            if bar:
                bar.update(len(chunk))
    return h.hexdigest()


# ------------------------------------------------------------- --scripts
# 這幾種呼叫代表「讀」或「寫」, 抓到就把它的第一個字串參數當路徑
READ_FN = {"read_csv", "read_excel", "read_parquet", "ExcelFile", "load_workbook"}
WRITE_FN = {"to_csv", "to_excel", "to_parquet", "ExcelWriter", "save"}
PATHY = re.compile(r"[\\/]|\.(csv|xlsx|xls|txt|parquet|json)\b", re.I)


def _lit(node) -> str | None:
    if isinstance(node, ast.Constant) and isinstance(node.value, str):
        return node.value
    if isinstance(node, ast.JoinedStr):       # f-string
        out = []
        for v in node.values:
            if isinstance(v, ast.Constant):
                out.append(str(v.value))
            else:
                out.append("{…}")
        return "".join(out)
    return None


def scan_script(p: Path) -> None:
    raw = p.read_text(encoding="utf-8", errors="replace")
    add("")
    add("-" * 72)
    add(f"  {p.name}   {len(raw):,} 字元   {raw.count(chr(10)) + 1:,} 行")
    add("-" * 72)
    try:
        tree = ast.parse(raw)
    except SyntaxError as e:
        add(f"    !! 當成 Python 解析失敗 (第 {e.lineno} 行: {e.msg})")
        add("    改用文字搜尋:")
        for i, ln in enumerate(raw.splitlines(), 1):
            if PATHY.search(ln) and ("=" in ln or "(" in ln):
                add(f"      第 {i} 行  {ln.strip()[:100]}")
        return

    imports = sorted({(a.name.split(".")[0] if isinstance(n, ast.Import)
                       else (n.module or "").split(".")[0])
                      for n in ast.walk(tree)
                      if isinstance(n, (ast.Import, ast.ImportFrom))
                      for a in (n.names if isinstance(n, ast.Import) else [n])})
    add(f"    import: {', '.join(x for x in imports if x)}")

    # 路徑寫死在哪 —— 要改的就是這幾行
    hard: list = []
    for n in ast.walk(tree):
        v = _lit(n)
        if v and PATHY.search(v) and len(v) > 3:
            hard.append((getattr(n, "lineno", 0), v))
    seen = set()
    if hard:
        add("    路徑字串 (要換環境就改這幾行):")
        for ln, v in sorted(hard):
            if v in seen:
                continue
            seen.add(v)
            add(f"      第 {ln:>4} 行  {v[:96]}")

    # 讀 / 寫
    rd, wr = [], []
    for n in ast.walk(tree):
        if not isinstance(n, ast.Call):
            continue
        fn = (n.func.attr if isinstance(n.func, ast.Attribute)
              else getattr(n.func, "id", ""))
        arg = _lit(n.args[0]) if n.args else None
        if fn in READ_FN:
            rd.append((getattr(n, "lineno", 0), fn, arg))
        elif fn in WRITE_FN:
            wr.append((getattr(n, "lineno", 0), fn, arg))
        elif fn == "open":
            mode = (_lit(n.args[1]) if len(n.args) > 1 else "r") or "r"
            (wr if "w" in mode or "a" in mode else rd).append(
                (getattr(n, "lineno", 0), f"open({mode})", arg))
        elif fn in ("glob", "iglob", "rglob"):
            rd.append((getattr(n, "lineno", 0), "glob", arg))
    for lab, items in (("讀", rd), ("寫", wr)):
        if items:
            add(f"    {lab}:")
            for ln, fn, arg in sorted(items):
                add(f"      第 {ln:>4} 行  {fn:<16}{(arg or '(變數, 非字面值)')[:70]}")

    # 頂層設定變數 —— 通常就是要改的東西
    # 頂層設定變數 —— 通常就是要改的東西。算出來的 (非字面值) 要把原始
    # 那一行印出來, 否則「BASE_FOLDER = (非字面值)」等於什麼都沒講。
    src = raw.splitlines()
    cfg = []
    for n in tree.body:
        if isinstance(n, ast.Assign) and len(n.targets) == 1:
            t = getattr(n.targets[0], "id", "")
            if re.search(r"folder|path|dir|file|input|output|src|dst|root",
                         t, re.I):
                v = _lit(n.value)
                if v is None:
                    seg = "\n".join(src[n.lineno - 1:getattr(n, "end_lineno",
                                                             n.lineno)])
                    v = " ".join(seg.split())
                cfg.append((n.lineno, t, v))
    if cfg:
        add("    頂層設定變數:")
        for ln, t, v in cfg:
            add(f"      第 {ln:>4} 行  {v[:110]}")

    fns = [n.name for n in tree.body if isinstance(n, ast.FunctionDef)]
    if fns:
        add(f"    函式: {', '.join(fns[:15])}"
            + (f" ... 共 {len(fns)} 個" if len(fns) > 15 else ""))
    has_main = any(isinstance(n, ast.If) and "__main__" in ast.dump(n.test)
                   for n in tree.body)
    add(f"    怎麼跑: {'python ' + p.stem + '.py' if has_main else '沒有 main guard, 整份由上而下執行'}")
    if any("input(" in ln for ln in raw.splitlines()):
        add("    !! 有 input() —— 用 subprocess 跑要先餵換行, 否則會卡住")


def mode_scripts(root: Path) -> None:
    d = root / MISC
    if not d.is_dir():
        d = root
    ps = [p for p in sorted(d.rglob("*"))
          if p.suffix.lower() in (".txt", ".py") and p.is_file()]
    add("")
    add("=" * 72)
    add(f"副檔名是 .txt 但內容是 Python 的檔 ({d})")
    add("=" * 72)
    if not ps:
        add("  找不到 .txt / .py")
        return
    for p in ps:
        scan_script(p)


# ------------------------------------------------------------- --dup
def mode_dup(root: Path, dirs: list[str], quick: bool) -> None:
    ds = [root / x for x in dirs] if dirs else [
        d for d in sorted(root.iterdir())
        if d.is_dir() and d.name != MISC and d.name != OUT_PARQUET
        and any(d.glob("*.csv"))]
    if len(ds) != 2:
        sys.exit(f"要正好兩個資料夾, 找到 {len(ds)} 個: "
                 + ", ".join(d.name for d in ds) + "\n  (用 --dirs a,b 指定)")
    A, B = ds
    fa = sorted(A.rglob("*.csv"))
    fb = sorted(B.rglob("*.csv"))
    add("")
    add("=" * 72)
    add("兩個資料夾的檔是不是同一批")
    add("=" * 72)
    for d, fs in ((A, fa), (B, fb)):
        tot = sum(p.stat().st_size for p in fs)
        add(f"  {d.name}   {len(fs)} 檔   {mb(tot)}")

    add("")
    add("  系列盤點")
    for d, fs in ((A, fa), (B, fb)):
        ser: dict = {}
        for p in fs:
            ser.setdefault(series_of(p.name), []).append(p)
        add(f"    {d.name}")
        for sn, ps in sorted(ser.items(), key=lambda kv: -len(kv[1])):
            sz = sum(q.stat().st_size for q in ps)
            add(f"      {len(ps):>4} 檔  {mb(sz):>10}   [{sn}]")

    # 先用大小配 —— 秒級, 而且大小不同就一定不是同一個檔
    by_sz: dict = {}
    for p in fa:
        by_sz.setdefault(p.stat().st_size, []).append(p)
    cand = [(p, q) for q in fb for p in by_sz.get(q.stat().st_size, [])]
    add("")
    add(f"  大小相同的組合: {len(cand)}")
    if quick:
        for p, q in cand[:60]:
            add(f"    {p.stat().st_size:>15,}   {p.name[:44]:<46}{q.name[:44]}")
        if len(cand) > 60:
            add(f"    ... 另外 {len(cand) - 60} 組")
        add("")
        add("  (--quick: 只比大小。拿掉它會再算 SHA256 確認內容真的一樣)")
        return

    same, diff = [], []
    tot = sum(p.stat().st_size + q.stat().st_size for p, q in cand)
    add(f"  算 SHA256 確認 ({mb(tot)} 要讀)…")
    t0 = time.time()
    for p, q in cand:
        (same if sha256(p) == sha256(q) else diff).append((p, q))
    add(f"  ({time.time() - t0:.0f}s)")
    add("")
    add(f"  內容完全相同 {len(same)} 組   大小相同但內容不同 {len(diff)} 組")
    for lab, xs in (("相同", same), ("不同", diff)):
        if not xs:
            continue
        add(f"    {lab}:")
        for p, q in xs[:60]:
            add(f"      {p.stat().st_size:>15,}   {p.name[:42]:<44}{q.name[:42]}")
        if len(xs) > 60:
            add(f"      ... 另外 {len(xs) - 60} 組")
    dup_sz = sum(p.stat().st_size for p, _ in same)
    add("")
    add(f"  重複佔用 {mb(dup_sz)} —— 刪掉其中一份就省這麼多")
    ua = [p for p in fa if p not in {x for x, _ in same}]
    ub = [q for q in fb if q not in {y for _, y in same}]
    add(f"  只在 {A.name} 有 {len(ua)} 檔 {mb(sum(p.stat().st_size for p in ua))}")
    add(f"  只在 {B.name} 有 {len(ub)} 檔 {mb(sum(p.stat().st_size for p in ub))}")


# ------------------------------------------------------------- --parquet
def mode_parquet(root: Path, dry: bool, only: str | None,
                 typed: bool = False) -> None:
    try:
        import pyarrow as pa
        import pyarrow.csv as pacsv
        import pyarrow.parquet as pq
    except ImportError:
        sys.exit("要先裝 pyarrow:  pip install pyarrow")

    fs = [p for p in sorted(root.rglob("*.csv"))
          if OUT_PARQUET not in p.parts
          and (not only or only.lower() in p.name.lower())]
    add("")
    add("=" * 72)
    add(("估算" if dry else "轉換") + f" parquet ({len(fs)} 個檔)"
        + ("" if dry else ("   型別: " + ("pyarrow 自動推斷 (--typed)"
                                         if typed else "全部存成文字 (保真)"))))
    add("=" * 72)
    if not fs:
        add("  找不到 CSV")
        return
    out = root / OUT_PARQUET
    src_tot = got_tot = 0
    bad_rows: list = []
    bar = tqdm(total=sum(q.stat().st_size for q in fs), unit="B",
               unit_scale=True, desc="  轉換" if not dry else "  估算",
               leave=False)
    for i, p in enumerate(fs, 1):
        sz = p.stat().st_size
        rel = p.relative_to(root)
        bar.update(sz)
        t0 = time.time()
        try:
            if dry:
                # 只讀前面幾萬列推估 —— 全轉一次太慢, 而且壓縮比在同一個檔
                # 內部相當穩定 (欄位型態與重複程度不會只在後半段改變)
                ro = pacsv.ReadOptions(block_size=1 << 23)
                tbl = next(pacsv.open_csv(p, read_options=ro).__iter__())
                tbl = pa.Table.from_batches([tbl])
                tmp = root / f"_probe_{i}.parquet"
                pq.write_table(tbl, tmp, compression="zstd")
                ratio = (tbl.nbytes / tmp.stat().st_size) if tmp.stat().st_size else 0
                est = int(sz / ratio) if ratio else 0
                tmp.unlink()
                got_tot += est
                pass
            else:
                dst = out / rel.with_suffix(".parquet")
                dst.parent.mkdir(parents=True, exist_ok=True)
                # 預設把每一欄都存成文字。這批資料的單號有前導零 ("00123…"),
                # 讓 pyarrow 自己推型別會把它當整數, 零就掉了, 而且不同檔
                # 推出不同型別, 合併時會炸。要當存檔用就必須逐字保真。
                co = None
                if not typed:
                    co = pacsv.ConvertOptions(
                        column_types={c: pa.string() for c in csv_header(p)})
                w = None
                nrow = 0
                for batch in pacsv.open_csv(
                        p, read_options=pacsv.ReadOptions(block_size=1 << 26),
                        convert_options=co):
                    t = pa.Table.from_batches([batch])
                    if not typed:
                        nostr = [f.name for f in t.schema if f.type != pa.string()]
                        if nostr:
                            raise ValueError(
                                "這幾欄沒被指定成文字, 會失真: "
                                + ", ".join(repr(x) for x in nostr[:5]))
                    nrow += t.num_rows
                    if w is None:
                        w = pq.ParquetWriter(dst, t.schema, compression="zstd")
                    w.write_table(t)
                if w:
                    w.close()
                got = dst.stat().st_size
                got_tot += got
                # 核對列數 —— 要刪原檔之前一定要確認一列都沒漏。
                # parquet 的列數在 metadata 裡, 讀它不用掃全檔。
                pn = pq.ParquetFile(dst).metadata.num_rows
                ok = "OK" if pn == nrow else f"!! CSV {nrow:,} / parquet {pn:,}"
                if pn != nrow:
                    bad_rows.append(rel.name)
                if pn != nrow:
                    add(f"  !! {rel.name[:50]}  {ok}")
        except Exception as e:
            add(f"  !! {rel.name[:50]}  {type(e).__name__}: {e}")
            bad_rows.append(f"{rel.name}  ({type(e).__name__})")
            # 出錯的檔多半已經寫了一半, 留著會讓人以為轉好了
            half = out / rel.with_suffix(".parquet")
            if not dry and half.exists():
                half.unlink()
                add("        (已刪掉寫到一半的 parquet)")
            continue
        src_tot += sz
    bar.close()
    add("")
    add(f"  原始 {mb(src_tot)}  ->  parquet {mb(got_tot)}"
        + (f"   ({src_tot / got_tot:.1f}x, 省 {mb(src_tot - got_tot)})"
           if got_tot else ""))
    if dry:
        add("  (--dry 只取樣推估。實際轉換請拿掉 --dry)")
    elif bad_rows:
        add(f"  !! 有 {len(bad_rows)} 個檔的列數對不上, **不要刪原檔**:")
        for x in bad_rows[:20]:
            add(f"       {x}")
    else:
        add("  所有檔的 CSV 列數與 parquet 列數一致, 原檔可以刪。")


def _scan_root(root: Path) -> tuple[Path, list]:
    """要掃哪些檔。

    有 data/ 就只掃 data/ —— 專案資料夾底下常常還有別的東西 (別的工作的
    來源、上一輪的產出、工具自己寫的 csv), 全樹掃進來算出的涵蓋範圍是錯的。
    沒有 data/ 才退回掃整棵樹 (還沒分資料夾的情況)。
    """
    d = root / DATA_DIR
    if d.is_dir():
        fs = sorted(d.glob("*.csv")) + sorted(d.glob("*.parquet"))
        return d, fs
    skip = {OUT_PARQUET, OUT_DIR, MISC}
    fs = [q for q in sorted(root.rglob("*.csv")) + sorted(root.rglob("*.parquet"))
          if not skip & set(q.parts)]
    return root, fs


def mode_coverage(root: Path) -> None:
    """每個系列涵蓋哪些月份, 哪裡有洞 —— 「跑的是不是完整的」就看這張表。"""
    where, fs = _scan_root(root)
    cov: dict = {}
    bad: list = []
    for p in fs:
        sn = series_of(p.name)
        if p.name.lower() in ("fx rate.csv", "station region mapping.csv"):
            continue
        span, gran = parse_span(p.name)
        if span is None:
            bad.append(p)
            continue
        for ym in span:
            cov.setdefault(sn, {}).setdefault(ym, []).append((p, gran))
    add("")
    add("=" * 72)
    add(f"涵蓋範圍 (每個系列哪些月份有、哪些沒有)   掃 {where.name}/  "
        f"{len(fs)} 個檔")
    add("=" * 72)
    for p in bad:
        add(f"  !! 檔名看不出期間: {p.name}")
    if not cov:
        add("  沒有可判讀的檔")
        return

    allym = sorted({ym for d in cov.values() for ym in d})
    sers = sorted(cov, key=lambda k: -len(cov[k]))
    add("")
    add(f"  整體範圍 {allym[0][0]}-{allym[0][1]:02d} ~ "
        f"{allym[-1][0]}-{allym[-1][1]:02d}   共 {len(allym)} 個月")
    add("")
    for sn in sers:
        d = cov[sn]
        ks = sorted(d)
        # 該系列自己的起訖之間有沒有缺月
        span_all = []
        y, m = ks[0]
        while (y, m) <= ks[-1]:
            span_all.append((y, m))
            y, m = (y + 1, 1) if m == 12 else (y, m + 1)
        miss = [x for x in span_all if x not in d]
        gran = Counter(g for v in d.values() for _, g in v)
        nfile = len({q for v in d.values() for q, _ in v})
        add(f"  [{sn}]   {nfile} 檔覆蓋 "
            f"{len(d)}/{len(span_all)} 個月   "
            f"{ks[0][0]}-{ks[0][1]:02d} ~ {ks[-1][0]}-{ks[-1][1]:02d}")
        add(f"      顆粒度: " + ", ".join(f"{k} {v}" for k, v in gran.most_common()))
        if miss:
            add(f"      !! 缺 {len(miss)} 個月: "
                + ", ".join(f"{a}-{b:02d}" for a, b in miss[:18])
                + (" …" if len(miss) > 18 else ""))
        # 同一個月被多個檔蓋到 (半月×2 是正常, 更多就要看)
        over = {ym: v for ym, v in d.items()
                if len(v) > 1 and not {g for _, g in v} <= {"上半月", "下半月"}}
        if over:
            add(f"      !! {len(over)} 個月被多個檔重複覆蓋, 合併會重算:")
            for ym, v in sorted(over.items())[:6]:
                add(f"          {ym[0]}-{ym[1]:02d}  " +
                    " / ".join(f"{q.name[:40]}({g})" for q, g in v))

    add("")
    add("  各月份被幾個系列覆蓋")
    for i, sn in enumerate(sers, 1):
        add(f"    S{i} = [{sn}]")
    add("")
    add(f"    {'月份':<10}" + "".join(f"S{i:<8}" for i in range(1, len(sers) + 1)))
    # 只標「系列自己的起訖範圍內卻沒有檔」的月份。跨系列比「齊不齊」沒有
    # 意義 —— 各系列本來就涵蓋不同期間 (有的中途才開始, 有的提早結束)。
    rng = {sn: (min(cov[sn]), max(cov[sn])) for sn in sers}
    holes = 0
    for ym in allym:
        cells, hole = [], False
        for sn in sers:
            k = len(cov[sn].get(ym, []))
            lo, hi = rng[sn]
            if not k and lo <= ym <= hi:
                cells.append("洞".ljust(8))
                hole = True
            else:
                cells.append(("-" if not k else str(k)).ljust(9))
        holes += hole
        add(f"    {ym[0]}-{ym[1]:02d}   " + "".join(cells)
            + ("  <- 有洞" if hole else ""))
    add("")
    if holes:
        add(f"  !! {holes} 個月在某個系列自己的範圍內卻沒有檔 (上表標「洞」)")
    else:
        add("  每個系列在自己的起訖範圍內都沒有缺月。")
    add('     ("-" 是該系列本來就不涵蓋那個月, 不是缺)')


def mode_schema(root: Path) -> None:
    """比對每個檔的欄位版本 —— 要合併處理, 欄位不一致就會出事。"""
    where, fs = _scan_root(root)
    add("")
    add("=" * 72)
    add(f"欄位版本 (只讀表頭, 秒級)   掃 {where.name}/   {len(fs)} 個檔")
    add("=" * 72)
    ser: dict = {}
    for p in fs:
        if p.suffix.lower() == ".parquet":
            import pyarrow.parquet as pq
            hdr = ",".join(pq.ParquetFile(p).schema_arrow.names).encode()
        else:
            with open(p, "rb") as f:
                hdr = f.readline().rstrip(b"\r\n")
        ser.setdefault(series_of(p.name), {}).setdefault(hdr, []).append(p)
    for sn, byh in sorted(ser.items(), key=lambda kv: -sum(len(v) for v in kv[1].values())):
        n_ = sum(len(v) for v in byh.values())
        add("")
        add(f"  [{sn}]   {n_} 檔")
        base = max(byh, key=lambda k: len(byh[k]))
        bc = [c.strip().strip('"') for c in
              base.decode("utf-8", "replace").lstrip("\ufeff").split(",")]
        for h, ps in sorted(byh.items(), key=lambda kv: -len(kv[1])):
            cc = [c.strip().strip('"') for c in
                  h.decode("utf-8", "replace").lstrip("\ufeff").split(",")]
            note = ""
            if h != base:
                ex = [c for c in cc if c not in bc]
                mi = [c for c in bc if c not in cc]
                note = (("  多: " + ", ".join(ex[:6])) if ex else "") + \
                       (("  少: " + ", ".join(mi[:6])) if mi else "")
                note = note or "  (欄名相同, 順序不同)"
            add(f"      {len(cc):>3} 欄 × {len(ps):>3} 檔{note}")
            for q in ps[:2]:
                add(f"          {q.name}")
            if len(ps) > 2:
                add(f"          ... 另外 {len(ps) - 2} 個")
    add("")
    add("  合併處理之前, 同一個系列內部的欄位版本要先統一,")
    add("  否則 concat 出來的表會多出一堆全空的欄。")


def mode_verify(root: Path, only: str | None) -> None:
    """轉完之後、刪原檔之前跑這個。

    比三件事: 列數、欄名、以及頭尾各 200 列逐格比對。列數抓得到截斷,
    逐格比對抓得到型別被轉掉 (最怕的就是單號 "00123…" 變成 123)。
    """
    try:
        import pyarrow.parquet as pq
    except ImportError:
        sys.exit("要先裝 pyarrow:  pip install pyarrow")
    out = root / OUT_PARQUET
    if not out.is_dir():
        sys.exit(f"找不到 {out}, 先跑 --parquet")
    add("")
    add("=" * 72)
    add("核對 parquet 與原始 CSV")
    add("=" * 72)
    ok = bad = miss = 0
    todo = sorted(out.rglob("*.parquet"))
    bar = tqdm(total=len(todo), unit="檔", desc="  核對", leave=False)
    for dst in todo:
        bar.update(1)
        rel = dst.relative_to(out).with_suffix(".csv")
        src = root / rel
        if only and only.lower() not in src.name.lower():
            continue
        if not src.exists():
            add(f"  !! 找不到原始 CSV: {rel}")
            miss += 1
            continue
        pf = pq.ParquetFile(dst)
        pn = pf.metadata.num_rows
        pcols = [f.name for f in pf.schema_arrow]

        hdr = csv_header(src)
        head, tail, n = [], [], 0
        with open(src, "rb") as f:
            f.readline()
            for line in f:
                if not line.strip():
                    continue
                n += 1
                if len(head) < 200:
                    head.append(line.rstrip(b"\r\n"))
                tail.append(line.rstrip(b"\r\n"))
                if len(tail) > 200:
                    tail.pop(0)

        prob = []
        if pn != n:
            prob.append(f"列數 CSV {n:,} / parquet {pn:,}")
        if pcols != hdr:
            prob.append(f"欄名不同 (CSV {len(hdr)} / parquet {len(pcols)})")
        else:
            t_head = pf.read_row_group(0).slice(0, len(head)).to_pylist()
            for i, (raw, row) in enumerate(zip(head, t_head)):
                got = [("" if row[c] is None else str(row[c])) for c in pcols]
                want = next(csv.reader([raw.decode("utf-8", "replace")]))
                if got != want:
                    d = [(pcols[k], want[k], got[k])
                         for k in range(min(len(got), len(want)))
                         if got[k] != want[k]]
                    prob.append(f"第 {i + 2} 列有 {len(d)} 格對不上"
                                + (f": {d[0][0]} CSV={d[0][1]!r} "
                                   f"parquet={d[0][2]!r}" if d else ""))
                    break
        if prob:
            bad += 1
            add(f"  !! {rel.name[:52]}")
            for x in prob:
                add(f"       {x}")
        else:
            ok += 1
    bar.close()

    # 反方向也要查: 有 CSV 卻沒有對應的 parquet。轉換失敗的檔根本不會進到
    # 上面的迴圈, 「全部通過」講的只是轉成功的那些 —— 照著去刪原檔, 失敗
    # 的那幾個就是唯一一份, 刪了就沒了。
    done = {q.relative_to(out).with_suffix(".csv") for q in todo}
    src_all = [q for q in root.rglob("*.csv")
               if OUT_PARQUET not in q.parts
               and (not only or only.lower() in q.name.lower())]
    nopq = [q for q in src_all if q.relative_to(root) not in done]

    # 傳輸斷在區塊邊界的檔, 大小會剛好是 1MiB 的整數倍。真實 CSV 落在這種
    # 數字上的機率約百萬分之一, 所以這個訊號很乾淨。斷在換行處的話還會轉
    # 換成功、逐格比對也過 —— 只有大小看得出來。
    trunc = [q for q in src_all
             if q.stat().st_size and q.stat().st_size % (1 << 20) == 0]

    add("")
    add(f"  通過 {ok}   有問題 {bad}   找不到原始 CSV {miss}"
        + (f"   沒有 parquet 的 CSV {len(nopq)}" if nopq else ""))
    if nopq:
        add(f"  !! 這 {len(nopq)} 個 CSV 沒有對應的 parquet (轉換時失敗了), "
            "**絕對不要刪**:")
        for q in nopq[:12]:
            add(f"       {q.name[:60]}")
        if len(nopq) > 12:
            add(f"       … 另外 {len(nopq) - 12} 個")
    if trunc:
        add(f"  !! 這 {len(trunc)} 個 CSV 的大小剛好是 1MiB 的整數倍 —— "
            "傳輸斷在區塊邊界的特徵, 檔案很可能不完整:")
        for q in trunc:
            add(f"       {q.name[:52]}  {q.stat().st_size / (1 << 20):.0f} MiB 整")
        add("       斷點剛好落在換行處的話, 轉換與逐格比對都會過, "
            "只有大小看得出來。要重新拿。")
    if bad or nopq or trunc:
        add("  !! **不要刪原檔**")
    elif ok:
        add("  全部通過 (列數、欄名、頭 200 列逐格) —— 原檔可以封存或刪除")



# ------------------------------------------------------- --patch
# 客戶那幾支腳本帶客戶資訊, 不能進 public repo。所以改動本身寫成定點替換:
# 程式碼進 git, 被改的檔不進。每一處只認一小段原文, 對不上就跳過並說明。
COMBINE_HINT = "combine"
ARA_HINT = "collect_tax_summary"
ADJ_HINT = "adjustment_summary"
RESUME_HINT = "resume_from_sqlite"   # 中斷之後接著跑的那支
DATA_DIR = "data"          # 所有來源檔放這裡
OUT_DIR = "output"         # 所有產出寫這裡 (不存在會自己建)
BAK = ".bak"
CMP_DIAG_CAP = 5000     # 工作表大於這個列數就不做逐欄診斷 (讀原值很慢)
MAX_SHOW = 8

PARQUET_READER = '''

def parquet_chunks(path):
    """分批讀 parquet, 產出的 DataFrame 跟 csv_chunks 一樣是全字串欄。

    parquet 存的是全文字 (轉檔時就指定好), 所以這裡不會有型別推斷, 票號的
    前導零不會掉。空值在 parquet 是 None、在 read_csv(dtype=str) 是 NaN,
    但下游 clean() 一律 fillna("") , 兩者等價。
    """
    import pyarrow.parquet as pq

    handle = pq.ParquetFile(path)
    for batch in handle.iter_batches(batch_size=CHUNK):
        frame = batch.to_pandas()
        # 一定要是 object dtype。新版 pandas 會把 arrow 字串欄做成
        # string[pyarrow], 那種欄的 .str.match() 走 pyarrow 的 RE2 引擎,
        # 而 RE2 不支援反向參照 (\\1)。這支腳本的票號樣式正好用到,
        # read_csv(dtype=str) 給的 object 欄走 Python re 沒事, 只有
        # parquet 這條路會炸。轉回 object 就跟 CSV 完全同一條路。
        cast = {c: frame[c].astype(object) for c in frame.columns
                if frame[c].dtype != object}
        if cast:
            frame = frame.assign(**cast)
        yield frame

'''

EDITS_COMBINE = [
    ('磁碟空間估算改用實測的每列位元組',
     '_pq.ParquetFile(path)',
     '    source_bytes = 0\n    for path in files:\n        try:\n            source_bytes += os.path.getsize(path)\n        except OSError:\n            pass\n',
     '    source_bytes = 0\n    for path in files:\n        try:\n            # parquet 是壓縮的, 檔案大小不等於處理時的資料量。這個估算\n            # 是拿來擋「跑到一半磁碟滿」的, 用壓縮後的大小會低估一個\n            # 數量級, 等於沒擋。metadata 裡的 total_byte_size 也不行\n            # (那是編碼後的, 字典編碼之後遠小於真正的文字量)。所以讀\n            # 一小塊實際量每列多少位元組, 再乘總列數。\n            if os.path.splitext(path)[1].lower() == ".parquet":\n                import pyarrow as _pa\n                import pyarrow.parquet as _pq\n                _f = _pq.ParquetFile(path)\n                _b = next(_f.iter_batches(batch_size=20000), None)\n                if _b is not None and _b.num_rows:\n                    _t = _pa.Table.from_batches([_b])\n                    source_bytes += int(\n                        _f.metadata.num_rows * (_t.nbytes / _t.num_rows)\n                    )\n            else:\n                source_bytes += os.path.getsize(path)\n        except Exception:\n            try:\n                source_bytes += os.path.getsize(path)\n            except OSError:\n                pass\n'),
    ('加一個只過濾畫面的 Tee 子類別',
     'class QuietTee(Tee):',
     'if __name__ == "__main__":',
     'class QuietTee(Tee):\n    """只過濾「寫到畫面」那一路; 寫到 log 的一行都不少。\n\n    一輪跑下來畫面會刷幾千行, 但真要回頭查是看 _Run_Log.txt。所以畫面只留:\n    區段標題、計數與金額、警告與錯誤, 其餘的環境資訊與逐檔逐塊流水帳全部\n    收起來, 讀檔改成一條原地更新的進度列 (總數從 "Number of source files\n    found: N" 自己抓, 不必動任何迴圈)。\n\n    不想要就把 QUIET_CONSOLE 設成 False, 行為完全回到原本的 Tee。\n    """\n\n    QUIET_CONSOLE = True\n\n    # 含這些字的一定印 —— 寧可多印也不要把問題藏起來\n    KEEP_ALWAYS = ("!!", "WARNING", "Warning", "ERROR", "Error", "error",\n                   "Traceback", "Press Enter", "NOT ", "not complete",\n                   "missing", "Missing")\n\n    # 環境資訊與逐檔逐塊流水帳\n    SKIP_ON_SCREEN = (\n        "  - ",\n        "Script file", "Working folder", "Full path: ",\n        "FX file found", "Region mapping file found",\n        "Previous output files",\n        "Loading FX", "FX Columns", "FX lookup loaded",\n        "Loading station region", "Region stations loaded",\n        "Work folder", "Source files    ", "Estimated need",\n        "Free right now", "There is enough room",\n        "Reading as ", "Detected encoding:",\n        "Loaded total rows:",\n        "Source file: ",\n        "Carrier-related original columns",\n        "FLIGHT_CARRIER found", "FLIGHT_CARRIER sample",\n        "Key columns mapped", "Nonblank FLIGHT_CARRIER",\n        "[\'",\n        # 這兩個區段的內容上面全濾掉了, 標題留著只是個空殼\n        \'SOURCE COLUMN MAPPING CHECK\',\n        \'FLIGHT CARRIER AFTER NORMALIZATION\',\n    )\n\n    def __init__(self, stream, log_handle):\n        Tee.__init__(self, stream, log_handle)\n        self._buf = ""\n        self._total = 0\n        self._seen = 0\n        self._bar = False\n        self._last_blank = True\n        self._pending = None      # 扣住的區段標題\n        self._sep = 0\n        self._expect_title = False\n\n    def _emit(self, line):\n        # 標題等到真的有內容要印才放出來 —— 底下整段都被濾掉時, 只留一個\n        # 光禿禿的標題比不印還礙眼。\n        if self._pending is not None:\n            title, self._pending = self._pending, None\n            self._emit_raw(title)\n        self._emit_raw(line)\n\n    def _emit_raw(self, line):\n        if self._bar:\n            self.stream.write("\\n")\n            self._bar = False\n        self.stream.write(line + "\\n")\n\n    def _one_line(self, line):\n        stripped = line.strip()\n\n        if line.startswith("Number of source files found:"):\n            try:\n                self._total = int(line.rsplit(":", 1)[1].strip())\n            except Exception:\n                self._total = 0\n\n        if line.startswith("Reading: "):\n            self._seen += 1\n            if self._pending is not None:\n                title, self._pending = self._pending, None\n                self._emit_raw(title)\n            self.stream.write(\n                "\\r  \\u8b80\\u53d6 [{}/{}] {:<58}".format(\n                    self._seen, self._total or "?", line[9:].strip()[:58])\n            )\n            self.stream.flush()\n            self._bar = True\n            self._last_blank = False\n            return\n\n        if any(k in line for k in self.KEEP_ALWAYS):\n            self._emit(line)\n            self._last_blank = False\n            return\n\n        # 區段是 分隔線/標題/分隔線 三行。分隔線只是裝飾, 標題先扣住。\n        if stripped and set(stripped) <= {"=", "-"} and len(stripped) > 8:\n            self._sep += 1\n            if self._sep % 2:\n                self._expect_title = True\n            return\n        if self._expect_title:\n            self._expect_title = False\n            # 標題本身也可能在跳過清單裡 (整段內容都被濾掉的那幾個)\n            if not line.startswith(self.SKIP_ON_SCREEN):\n                self._pending = line\n            return\n\n        if line.startswith(self.SKIP_ON_SCREEN):\n            return\n\n        if not stripped:\n            if self._last_blank or self._pending is not None:\n                return            # 連續空行收成一行; 標題還扣著就不算內容\n            self._last_blank = True\n            self._emit_raw("")\n            return\n\n        self._last_blank = False\n        self._emit(line)\n\n    def write(self, text):\n        try:\n            self.log_handle.write(text)\n            self.log_handle.flush()\n        except Exception:\n            pass\n        if not self.QUIET_CONSOLE:\n            self.stream.write(text)\n            return len(text)\n        self._buf += text\n        while "\\n" in self._buf:\n            line, self._buf = self._buf.split("\\n", 1)\n            self._one_line(line)\n        return len(text)\n\n    def flush(self):\n        if self.QUIET_CONSOLE and self._buf:\n            self._one_line(self._buf)\n            self._buf = ""\n        self.stream.flush()\n        try:\n            self.log_handle.flush()\n        except Exception:\n            pass\n\n\nif __name__ == "__main__":'),
    ('畫面改用它 (log 不受影響)',
     'sys.stdout = QuietTee(original_stdout, log_handle)',
     'sys.stdout = Tee(original_stdout, log_handle)',
     'sys.stdout = QuietTee(original_stdout, log_handle)'),
    ('票號樣式改寫 (反向參照 -> RE2 也吃得下)',
     'r"|^0{10,}$|^1{10,}$',
     'SUSPICIOUS_TKT_PATTERN = (\n    r"^0*$|^0*1$|^0*2$|^0*3$|^0*4$|^0*5$|^0*6$|^0*7$|^0*8$|^0*9$"\n    r"|^(\\d)\\1{9,}$"\n    r"|^\\d*9{6,}$"\n    r"|^1234567890\\d*$"\n)',
     'SUSPICIOUS_TKT_PATTERN = (\n    r"^0*$|^0*1$|^0*2$|^0*3$|^0*4$|^0*5$|^0*6$|^0*7$|^0*8$|^0*9$"\n    # 原本這一段是  r"|^(\\d)\\1{9,}$"  —— 反向參照, 意思是「同一個數字\n    # 連續 10 個以上」。新版 pandas 的字串欄是 arrow 後端, .str.match()\n    # 會走 RE2, 而 RE2 不支援反向參照, 直接丟 ArrowInvalid。改寫成逐個\n    # 數字列舉, 判定結果完全相同 (8,016 個字串實測零差異), 不用反向參照。\n    r"|^0{10,}$|^1{10,}$|^2{10,}$|^3{10,}$|^4{10,}$"\n    r"|^5{10,}$|^6{10,}$|^7{10,}$|^8{10,}$|^9{10,}$"\n    r"|^\\d*9{6,}$"\n    r"|^1234567890\\d*$"\n)'),
    ("來源副檔名加 .parquet",
     'PARQUET_EXTENSIONS = (".parquet",)',
     'SOURCE_EXTENSIONS = (".csv",) + EXCEL_EXTENSIONS',
     'PARQUET_EXTENSIONS = (".parquet",)\n\n'
     'SOURCE_EXTENSIONS = (".csv",) + PARQUET_EXTENSIONS + EXCEL_EXTENSIONS'),
    ("插入 parquet_chunks",
     "def parquet_chunks(path):",
     "def source_chunks(path, bad_lines):",
     PARQUET_READER.strip("\n") + "\n\n\ndef source_chunks(path, bad_lines):"),
    ("dispatch 認得 .parquet",
     "elif extension in PARQUET_EXTENSIONS:",
     '''    if os.path.splitext(path)[1].lower() in EXCEL_EXTENSIONS:
        print("Reading as Excel workbook.")
        yield from excel_chunks(path)
    else:
        yield from csv_chunks(path, bad_lines)''',
     '''    extension = os.path.splitext(path)[1].lower()
    if extension in EXCEL_EXTENSIONS:
        print("Reading as Excel workbook.")
        yield from excel_chunks(path)
    elif extension in PARQUET_EXTENSIONS:
        print("Reading as Parquet file.")
        yield from parquet_chunks(path)
    else:
        yield from csv_chunks(path, bad_lines)'''),
    ("peek_header 認得 parquet",
     "return list(pq.ParquetFile(path).schema_arrow.names)",
     "    try:\n        if extension in EXCEL_EXTENSIONS:",
     "    try:\n"
     "        if extension in PARQUET_EXTENSIONS:\n"
     "            import pyarrow.parquet as pq\n"
     "            return list(pq.ParquetFile(path).schema_arrow.names)\n"
     "        if extension in EXCEL_EXTENSIONS:"),
]

EDITS_ARA = [
    ("找檔時也收 .parquet",
     'for pattern in ("*.csv", "*.parquet")',
     '''    files = sorted(
        path
        for path in input_folder.glob("*.csv")
        if path.is_file()
        and path.name.lower() not in OUTPUT_NAMES
    )''',
     '''    files = sorted(
        path
        for pattern in ("*.csv", "*.parquet")
        for path in input_folder.glob(pattern)
        if path.is_file()
        and path.name.lower() not in OUTPUT_NAMES
    )'''),
    ("排除兩個 lookup 檔, 不要當成來源資料讀",
     "_LOOKUP_NAMES = {",
     "OUTPUT_NAMES = {",
     "_LOOKUP_NAMES = {\n"
     '    "fx rate.csv",\n'
     '    "station region mapping.csv",\n'
     "}\n\n"
     "OUTPUT_NAMES = _LOOKUP_NAMES | {"),
    ("格式偵測: parquet 不用猜編碼",
     'return "parquet", None',
     '''def detect_csv_format(path):
    raw = path.read_bytes()[:65536]''',
     '''def detect_csv_format(path):
    # parquet 自帶 schema, 沒有編碼與分隔符可言, 用一個哨兵值往下傳
    if path.suffix.lower() == ".parquet":
        return "parquet", None

    raw = path.read_bytes()[:65536]'''),
    ("讀表頭: parquet 走 schema",
     "pq.ParquetFile(path).schema_arrow.names",
     '''            header = pd.read_csv(
                path,
                sep=separator,
                encoding=encoding,
                nrows=0,
            ).columns''',
     '''            if encoding == "parquet":
                import pyarrow.parquet as pq
                header = pd.Index(
                    pq.ParquetFile(path).schema_arrow.names
                )
            else:
                header = pd.read_csv(
                    path,
                    sep=separator,
                    encoding=encoding,
                    nrows=0,
                ).columns'''),
    ("分批讀: parquet 走 iter_batches",
     "def _parquet_reader(_path=path):",
     '''                chunk_reader = pd.read_csv(
                    path,
                    sep=separator,
                    encoding=encoding,
                    dtype=str,
                    keep_default_na=False,
                    chunksize=CHUNK_SIZE,
                    low_memory=False,
                )''',
     '''                if encoding == "parquet":
                    import pyarrow.parquet as pq

                    def _parquet_reader(_path=path):
                        handle = pq.ParquetFile(_path)
                        for batch in handle.iter_batches(
                            batch_size=CHUNK_SIZE
                        ):
                            frame = batch.to_pandas().fillna("")
                            cast = {
                                c: frame[c].astype(object)
                                for c in frame.columns
                                if frame[c].dtype != object
                            }
                            if cast:
                                frame = frame.assign(**cast)
                            yield frame

                    chunk_reader = _parquet_reader()
                else:
                    chunk_reader = pd.read_csv(
                        path,
                        sep=separator,
                        encoding=encoding,
                        dtype=str,
                        keep_default_na=False,
                        chunksize=CHUNK_SIZE,
                        low_memory=False,
                    )'''),
]


# ---------------------------------------------------------------- --paths
# 把「輸入」與「輸出」分到兩個固定資料夾: <腳本旁>/data 與 <腳本旁>/output。
# 原本三支都是「跟腳本同一層」既讀又寫, 來源與產出混在一起, 重跑一次就可能
# 把上一輪的產出當成輸入。分開之後也方便把整個 output/ 交出去比對。
EDITS_COMBINE_PATHS = [
    ("輸入改 data/, 輸出改 output/",
     'OUTPUT_FOLDER = SCRIPT_PATH.parent / "' + OUT_DIR + '"',
     "BASE_FOLDER = SCRIPT_PATH.parent",
     'BASE_FOLDER = SCRIPT_PATH.parent / "' + DATA_DIR + '"\n\n'
     'OUTPUT_FOLDER = SCRIPT_PATH.parent / "' + OUT_DIR + '"'),

    ("建 output/ 並取字串路徑",
     "output_folder = str(OUTPUT_FOLDER)",
     "    base_folder = str(BASE_FOLDER)",
     "    base_folder = str(BASE_FOLDER)\n"
     "    output_folder = str(OUTPUT_FOLDER)\n"
     "    os.makedirs(output_folder, exist_ok=True)"),

    ("兩個 Excel 產出寫到 output/",
     '"excel_output": os.path.join(output_folder, EXCEL_OUTPUT_NAME)',
     '        "excel_output": os.path.join(base_folder, EXCEL_OUTPUT_NAME),\n'
     '        "detail_excel_output": os.path.join(\n'
     "            base_folder, DETAIL_EXCEL_OUTPUT_NAME\n"
     "        ),",
     '        "excel_output": os.path.join(output_folder, EXCEL_OUTPUT_NAME),\n'
     '        "detail_excel_output": os.path.join(\n'
     "            output_folder, DETAIL_EXCEL_OUTPUT_NAME\n"
     "        ),"),

    ("執行紀錄寫到 output/",
     "log_path = os.path.join(str(OUTPUT_FOLDER), RUN_LOG_NAME)",
     "log_path = os.path.join(str(BASE_FOLDER), RUN_LOG_NAME)",
     "log_path = os.path.join(str(OUTPUT_FOLDER), RUN_LOG_NAME)"),
]

EDITS_ARA_PATHS = [
    ("輸出預設改 output/ (順便拿掉寫死的絕對路徑)",
     'return Path(__file__).resolve().parent / "' + OUT_DIR + '"',
     "def default_output_folder():\n    return Path(",
     "def default_output_folder():\n"
     '    return Path(__file__).resolve().parent / "' + OUT_DIR + '"\n\n\n'
     "def _unused_default_output_folder():\n    return Path("),

    ("沒給參數時預設讀 data/, 不彈視窗",
     "_default_input = Path(__file__).resolve().parent",
     "    else:\n        input_folder = select_folder()",
     "    else:\n"
     '        _default_input = Path(__file__).resolve().parent / "'
     + DATA_DIR + '"\n'
     "        input_folder = (\n"
     "            _default_input.resolve()\n"
     "            if _default_input.is_dir()\n"
     "            else select_folder()\n"
     "        )"),
]

EDITS_ADJ_PATHS = [
    ("讀寫都改 output/ (它讀的是上一支的產出)",
     'SOURCE_FOLDER = Path(__file__).resolve().parent / "' + OUT_DIR + '"',
     "    SOURCE_FOLDER = Path(__file__).resolve().parent\n"
     "except NameError:\n"
     "    SOURCE_FOLDER = Path.cwd()",
     '    SOURCE_FOLDER = Path(__file__).resolve().parent / "'
     + OUT_DIR + '"\n'
     "except NameError:\n"
     '    SOURCE_FOLDER = Path.cwd() / "' + OUT_DIR + '"\n'
     "SOURCE_FOLDER.mkdir(parents=True, exist_ok=True)"),
]


def _patch_find(root: Path, hint: str) -> list[Path]:
    return [p for p in sorted(root.glob("*.py"))
            if hint in p.name.lower() and p.name != Path(__file__).name]


def _patch_apply(p: Path, edits: list, dry: bool) -> bool:
    text = p.read_text(encoding="utf-8", errors="replace")
    done = skipped = 0
    fail = []
    for label, marker, old, new in edits:
        # 用一段「只有改過才會出現」的標記判斷, 不要用 new 的第一行 ——
        # 有些編輯是在原文前面插入東西, 原文改完之後還在, 會誤判成沒改過。
        if marker in text:
            skipped += 1
            continue
        if old not in text:
            fail.append(label)
            skipped += 1
            continue
        text = text.replace(old, new, 1)
        done += 1
    for lab in fail:
        add(f"  !! {p.name[:44]}  這一處找不到要替換的原文, 沒有改: {lab}")
        add("     (腳本版本可能跟預期不同)")
    if not done:
        add(f"  - {p.name[:52]}   已經是改好的狀態")
        return False
    if dry:
        add(f"  + {p.name[:52]}   會改 {done} 處"
            + (f", {len(fail)} 處找不到" if fail else ""))
        return False
    bak = p.with_suffix(p.suffix + BAK)
    if not bak.exists():
        bak.write_text(p.read_text(encoding="utf-8", errors="replace"),
                       encoding="utf-8")
    p.write_text(text, encoding="utf-8")
    # 改完先確認還是合法的 Python, 免得把一支 4,700 行的腳本改壞了才發現
    import ast
    try:
        ast.parse(text)
        add(f"  + {p.name[:52]}   改了 {done} 處, 語法 OK, 備份 {bak.name}")
    except SyntaxError as e:
        add(f"  !! {p.name}  改完語法錯了 (第 {e.lineno} 行: {e.msg}), 已還原")
        p.write_text(bak.read_text(encoding="utf-8"), encoding="utf-8")
        return False
    return True



# 換引擎只加一行 import。connect 的呼叫可能跨好幾行、帶一堆 sqlite 專屬參數
# (timeout、check_same_thread…), 去對準那個呼叫又脆又容易改錯; 改 import 之後
# 整個 sqlite3 命名空間都由 duck_shim 接手, 呼叫端一個字都不用動。
#
# 注意是「插入一行」不是「取代」: 原本可能寫成 `import sqlite3, os, sys`, 在
# 那個 import 上做字串取代會把後面幾個模組一起吃掉。插在整行之後, 後綁定的
# 名字勝出, 原本那行照樣執行, 不管它長什麼樣都不會壞。
DUCK_LINE = "import duck_shim as sqlite3   # 引擎換 DuckDB, 差異收在 duck_shim.py"
DUCK_IMPORT = re.compile(r"^(\s*)(?:import\s+sqlite3\b|from\s+sqlite3\s+import\b)")


def _unpatch_duck(p: Path, dry: bool) -> bool:
    """只拿掉引擎那一行, 其他修改一律保留。

    --revert 是把腳本還原成完全原始的狀態, 連「工作資料夾指向 data/」那個
    修改也會一起消失, 於是腳本跑去掃根目錄。要 A/B 比兩個引擎時需要的是
    只換引擎、其他不動。
    """
    text = p.read_text(encoding="utf-8", errors="replace")
    if "duck_shim" not in text:
        add(f"  - {p.name[:52]}   本來就是 sqlite")
        return False
    keep = [ln for ln in text.splitlines(keepends=True)
            if "duck_shim" not in ln]
    if dry:
        add(f"  + {p.name[:52]}   會拿掉 {len(text.splitlines()) - len(keep)} 行")
        return False
    out = "".join(keep)
    try:
        ast.parse(out)
    except SyntaxError as e:
        add(f"  !! {p.name[:44]}  拿掉之後語法壞了 ({e}), 沒有改")
        return False
    p.write_text(out, encoding="utf-8")
    add(f"  - {p.name[:52]}   引擎換回 sqlite (其他修改保留)")
    return True


def _patch_duck(p: Path, dry: bool) -> bool:
    text = p.read_text(encoding="utf-8", errors="replace")
    if "duck_shim" in text:
        add(f"  - {p.name[:52]}   已經是改好的狀態")
        return False
    lines = text.splitlines(keepends=True)
    at = None
    for i, ln in enumerate(lines):
        if DUCK_IMPORT.match(ln):
            at = i
            indent = DUCK_IMPORT.match(ln).group(1)
            break
    if at is None:
        add(f"  !! {p.name[:44]}  找不到 import sqlite3, 沒有改")
        return False
    if dry:
        add(f"  + {p.name[:52]}   會在第 {at + 1} 行之後插入一行")
        return False
    bak = p.with_suffix(p.suffix + BAK)
    if not bak.exists():
        bak.write_text(text, encoding="utf-8")
    lines.insert(at + 1, f"{indent}{DUCK_LINE}\n")
    out = "".join(lines)
    p.write_text(out, encoding="utf-8")
    import ast
    try:
        ast.parse(out)
        add(f"  + {p.name[:52]}   第 {at + 2} 行插入 duck_shim, 語法 OK, "
            f"備份 {bak.name}")
    except SyntaxError as e:
        add(f"  !! {p.name}  改完語法錯了 (第 {e.lineno} 行: {e.msg}), 已還原")
        p.write_text(bak.read_text(encoding="utf-8"), encoding="utf-8")
        return False
    return True


def mode_patch(root: Path, dry: bool, paths: bool, revert: bool,
               duck: bool = False, unduck: bool = False) -> None:
    add("")
    add("=" * 72)
    add("改客戶腳本" + (" (--dry: 只看不寫)" if dry else ""))
    add("=" * 72)
    if revert:
        n = 0
        for bak in sorted(root.glob(f"*.py{BAK}")):
            tgt = bak.with_suffix("")
            tgt.write_text(bak.read_text(encoding="utf-8"), encoding="utf-8")
            add(f"  還原 {tgt.name}")
            n += 1
        add(f"  共還原 {n} 個檔" if n else "  找不到 .bak")
        if n:
            add("")
            add("  ! --revert 是還原成**完全原始**的狀態, 連「工作資料夾指向")
            add("    data/」那個修改也一起沒了, 腳本會跑去掃根目錄。")
            add("    要接著跑的話先: python csvtools.py --patch --paths")
            add("    只想換引擎不要動別的: python csvtools.py --patch --sqlite")
        return
    groups = [(COMBINE_HINT, EDITS_COMBINE, "合併/重複偵測"),
              (ARA_HINT, EDITS_ARA, "月度彙總")]
    if unduck:
        hit = 0
        for hint in (COMBINE_HINT, ARA_HINT, RESUME_HINT):
            for q in _patch_find(root, hint):
                hit += _unpatch_duck(q, dry)
        add("")
        add(f"  換回 sqlite {hit} 個檔。路徑等其他修改沒有動。")
        add("  要換回 DuckDB: python csvtools.py --patch --duck")
        return
    if duck:
        if not (root / "duck_shim.py").exists():
            sys.exit(f"找不到 {root / 'duck_shim.py'}, 先把它放到腳本旁邊")
        hit = 0
        for hint, what in ((COMBINE_HINT, "合併/重複偵測"),
                           (ARA_HINT, "月度彙總"),
                           (RESUME_HINT, "中斷續跑")):
            ps = _patch_find(root, hint)
            if not ps:
                if hint != RESUME_HINT:      # 這支不一定存在, 沒有就算了
                    add(f"\n  找不到{what}腳本 (檔名要含 {hint!r})")
                continue
            for q in ps:
                hit += _patch_duck(q, dry)
        add("")
        if dry:
            add("  --dry 結束, 沒有改任何檔。")
        else:
            add(f"  改好 {hit} 個檔。原檔備份在 *.py{BAK}, 要還原就加 --revert")
        add("  duck_shim 會吞掉 PRAGMA、拿掉 WITHOUT ROWID、補上 executescript,")
        add("  並把 NULL 排序設成跟 sqlite 一樣 "
            "(nulls_first_on_asc_last_on_desc)。")
        add("  客戶腳本的 SQL 一句沒動。")
        return
    if paths:
        groups = [
            (COMBINE_HINT, EDITS_COMBINE + EDITS_COMBINE_PATHS, "合併/重複偵測"),
            (ARA_HINT, EDITS_ARA + EDITS_ARA_PATHS, "月度彙總"),
            (ADJ_HINT, EDITS_ADJ_PATHS, "案件彙總"),
        ]
    hit = 0
    for hint, edits, what in groups:
        ps = _patch_find(root, hint)
        if not ps:
            add(f"\n  找不到{what}腳本 (檔名要含 {hint!r})")
            continue
        for p in ps:
            hit += _patch_apply(p, edits, dry)
    add("")
    if dry:
        add("  --dry 結束, 沒有改任何檔。")
    else:
        add(f"  改好 {hit} 個檔。原檔備份在 *.py{BAK}, "
            "要還原就加 --revert")
    add("  ⚠️ parquet 必須是全文字版本 (--parquet 預設), 否則前導零已經掉了")


# ------------------------------------------------------- --cmp
def stem_key(name: str) -> str:
    """檔名去掉「括號註記」與「日期尾碼」之後當配對鍵。

    對方的檔名帶了執行情境與日期 (例如 "…(某某範圍)-20260911"),
    我方是乾淨的 "Output_analysis"。要比內容就得先認出這是同一份東西。
    """
    n = Path(name).stem
    n = re.sub(r"\(.*?\)", " ", n)              # 括號裡的範圍註記
    n = re.sub(r"[-_]?20\d{6}.*$", " ", n)      # -20260911 之後全丟
    n = re.sub(r"[-_]?\d{4}-\d{2}.*$", " ", n)
    n = re.sub(r"[_\-]+", " ", n)
    return " ".join(n.split()).lower()


_CMP_LOOSE = False       # --cmp --loose: "|" 串起來的多值當成無序集合

# 浮點加法不符合結合律: sqlite 逐列累加, DuckDB 向量化平行累加, 同一批數字
# 的 SUM() 會差最後幾個位元 (實測 6,106.09 對 6,106.090000000001)。這是累加
# 順序造成的, 不是邏輯差異, 而且無法消除 —— 要完全一致就得強迫單執行緒逐列
# 相加, 那等於放棄換引擎的意義。
#
# 所以數字取 12 位有效位數再比, 但**必須留下痕跡**: 只要有格子是靠這個容差
# 才對上的, 就記下來, 報告最後會講有幾個、最大相對差多少。稽核底稿要寫得出
# 「差異止於 1e-N, 屬浮點累加順序」, 不能讓它悄悄消失。
_CMP_SIG = 12
_NEAR = {"n": 0, "max": 0.0, "multi": 0}


def _numfmt(f: float) -> str:
    if f == int(f):
        return str(int(f))
    g = float(f"{f:.{_CMP_SIG}g}")
    if g != f:
        _NEAR["n"] += 1
        d = abs(f - g) / max(abs(f), 1e-30)
        _NEAR["max"] = max(_NEAR["max"], d)
    return repr(g)


def _norm(v) -> str:
    """把一個格子正規化成可比對的形式。

    同一支腳本、同一批資料, 從 CSV 讀進來是數字、從 parquet (全文字) 讀進
    來是字串, 寫進 Excel 之後一邊是 120 一邊是 "120"。逐字比會整份比不中,
    但那是型別差異不是資料差異。數字一律化成同一種寫法, 日期化成 ISO。
    """
    if v is None:
        return ""
    if isinstance(v, bool):
        return "1" if v else "0"
    if isinstance(v, (dt.datetime, dt.date)):
        return v.isoformat()
    if isinstance(v, (int, float)):
        return _numfmt(float(v))
    t = str(v).strip()
    if not t:
        return ""
    if _CMP_LOOSE and ("|" in t or "," in t):
        # GROUP_CONCAT 沒有 ORDER BY 時元素順序是未定義的, 換引擎就會不一樣。
        # 內容一樣只是排列不同, 不該算差異。
        #
        # 分隔符兩種都要認: 腳本有的地方用預設的 ",", 有的地方指定 " | "。
        # 而且一定要 strip —— " | " 拆開之後第一個 token 沒有前導空格、其餘
        # 有, 不去掉空格的話排序結果會跟著元素位置變, 等於沒排。
        sep = "|" if "|" in t else ","
        parts = sorted(x.strip() for x in t.split(sep))
        out = sep.join(parts)
        if out != t:
            _NEAR["multi"] += 1
        return out
    try:
        f = float(t)
    except ValueError:
        return t
    # "1,234" 這種帶千分位的不要當數字, 不然會跟 1234 混在一起
    if any(c in t for c in ",%") or t.lower() in ("nan", "inf", "-inf"):
        return t
    return _numfmt(f)


def _h(row) -> bytes:
    flat = "\x1f".join(_norm(v) for v in row)
    return hashlib.blake2b(flat.encode("utf-8", "replace"),
                           digest_size=8).digest()


def sheet_rows(p: Path, name: str, cap: int) -> list:
    """把某一張工作表的列讀成 [值list]。給逐欄診斷用, 所以要原值不是雜湊。"""
    out: list = []
    if p.suffix.lower() == ".csv":
        with open(p, encoding="utf-8-sig", newline="", errors="replace") as f:
            rd = csv.reader(f)
            next(rd, None)
            for row in rd:
                if any(x.strip() for x in row):
                    out.append([_norm(v) for v in row])
                    if len(out) >= cap:
                        break
        return out
    from openpyxl import load_workbook
    wb = load_workbook(p, read_only=True, data_only=True)
    try:
        if name not in wb.sheetnames:
            return out
        it = wb[name].iter_rows(values_only=True)
        next(it, None)
        for row in it:
            if row is None or all(v is None for v in row):
                continue
            out.append([_norm(v) for v in row])
            if len(out) >= cap:
                break
    finally:
        wb.close()
    return out


def sheets_of(p: Path) -> dict:
    """回 {工作表名: (欄名list, 列數, 行雜湊Counter)}。CSV 當成單一張表。"""
    out: dict = {}
    if p.suffix.lower() == ".csv":
        with open(p, encoding="utf-8-sig", newline="", errors="replace") as f:
            rd = csv.reader(f)
            hdr = next(rd, [])
            c: Counter = Counter()
            n = 0
            for row in rd:
                if any(x.strip() for x in row):
                    c[_h(row)] += 1
                    n += 1
        out["(csv)"] = (hdr, n, c)
        return out

    from openpyxl import load_workbook
    wb = load_workbook(p, read_only=True, data_only=True)
    try:
        for ws in wb.worksheets:
            it = ws.iter_rows(values_only=True)
            hdr = ["" if v is None else str(v) for v in next(it, ())]
            c = Counter()
            n = 0
            for row in it:
                if row is None or all(v is None for v in row):
                    continue
                c[_h(row)] += 1
                n += 1
            out[ws.title] = (hdr, n, c)
    finally:
        wb.close()
    return out


def _cmp_pair(a: Path, b: Path, quick: bool) -> list[str]:
    """回「問題清單」。空清單 = 兩邊一致。"""
    probs: list[str] = []
    sa, sb = sheets_of(a), sheets_of(b)
    only_a = [k for k in sa if k not in sb]
    only_b = [k for k in sb if k not in sa]
    if only_a:
        probs.append(f"只有我方有的工作表 ({len(only_a)}): "
                     + ", ".join(only_a[:MAX_SHOW]))
    if only_b:
        probs.append(f"只有對方有的工作表 ({len(only_b)}): "
                     + ", ".join(only_b[:MAX_SHOW]))

    for k in [x for x in sa if x in sb]:
        ha, na, ca = sa[k]
        hb, nb, cb = sb[k]
        if ha != hb:
            xa = [c for c in ha if c not in hb]
            xb = [c for c in hb if c not in ha]
            msg = f"[{k}] 欄不同 ({len(ha)} / {len(hb)} 欄)"
            if xa:
                msg += "  只有我方: " + ", ".join(xa[:6])
            if xb:
                msg += "  只有對方: " + ", ".join(xb[:6])
            probs.append(msg)
            continue
        if na != nb:
            probs.append(f"[{k}] 列數不同  我方 {na:,} / 對方 {nb:,}"
                         f"  (差 {abs(na - nb):,})")
        if quick:
            continue
        d = ca - cb
        e = cb - ca
        oa, ob = sum(d.values()), sum(e.values())
        if oa or ob:
            probs.append(f"[{k}] 內容不同  只在我方 {oa:,} 列 / "
                         f"只在對方 {ob:,} 列  (共同 {na - oa:,})")
            # 「列數一樣但一列都對不上」通常是某一欄整欄不同, 不是資料不同。
            # 每次拿掉一欄重算, 能把差異解釋光的那一欄就是元兇。工作表太大
            # 就跳過 —— 讀原值很慢, 而且小表已經足夠指認出是哪一欄。
            if ha == hb and 0 < max(na, nb) <= CMP_DIAG_CAP:
                ra_ = sheet_rows(a, k, CMP_DIAG_CAP)
                rb_ = sheet_rows(b, k, CMP_DIAG_CAP)
                w = min(len(ha), min((len(r) for r in ra_ + rb_), default=0))
                if w:
                    hits = fc_explain(ha[:w], [r[:w] for r in ra_],
                                      [r[:w] for r in rb_])
                    if hits:
                        n_, _i, c_, va, vb = hits[0]
                        probs.append(
                            f"    拿掉 {fc_cname(c_)} 之後 {n_:,}/"
                            f"{min(len(ra_), len(rb_)):,} 列對上"
                            f"   我方 {va[:30]!r}  對方 {vb[:30]!r}")
                    else:
                        probs.append("    逐欄診斷: 任何單一欄都解釋不了, "
                                     "差異橫跨多欄")
    return probs



def mode_cmp(root: Path, ours: str, theirs: str, only: str | None,
             quick: bool, loose: bool = False) -> None:
    global _CMP_LOOSE
    _CMP_LOOSE = loose
    A, B = root / ours, root / theirs
    for d in (A, B):
        if not d.is_dir():
            add(f"  !! 找不到 {d}")
            return

    def pick(d: Path) -> list:
        return [p for p in sorted(d.iterdir())
                if p.suffix.lower() in (".xlsx", ".xlsm", ".csv")
                and not p.name.startswith("~$")]

    fa_, fb_ = pick(A), pick(B)
    add("")
    add("=" * 72)
    add(f"產出比對   我方 {ours} ({len(fa_)} 檔)   對方 {theirs} ({len(fb_)} 檔)"
        + ("   --loose: 多值欄位當成無序" if loose else ""))
    add("=" * 72)
    by_b: dict = {}
    for p in fb_:
        by_b.setdefault(stem_key(p.name), []).append(p)
    pairs, solo_a = [], []
    for p in fa_:
        m = by_b.get(stem_key(p.name), [])
        if m:
            pairs.extend((p, q) for q in m)
        else:
            solo_a.append(p)
    solo_b = [q for q in fb_ if q not in {y for _, y in pairs}]
    # --only 要在配對之後才篩 —— 兩邊的檔名本來就不一樣 (我方是乾淨的
    # Output_analysis.xlsx, 對方帶了執行情境與日期), 分別去篩會把我方整個
    # 篩光, 變成「配對到 0 組」。篩的是配對, 只要有一邊對得上就留。
    if only:
        k = only.lower()
        pairs = [(a, b) for a, b in pairs
                 if k in a.name.lower() or k in b.name.lower()]
        add(f"  --only {only!r}: 只比對得上的 {len(pairs)} 組")
    add(f"  配對到 {len(pairs)} 組")
    for p, q in pairs:
        add(f"    {p.name[:40]:<42}<->  {q.name[:44]}")
    if solo_a:
        add(f"  只有我方有 {len(solo_a)} 個: "
            + ", ".join(x.name[:40] for x in solo_a[:MAX_SHOW]))
    if solo_b:
        add(f"  只有對方有 {len(solo_b)} 個: "
            + ", ".join(x.name[:40] for x in solo_b[:MAX_SHOW]))
    add("")
    same = diff = err = 0
    bar = tqdm(total=len(pairs), desc="  比對", unit="組", leave=False)
    for p, q in pairs:
        bar.update(1)
        try:
            probs = _cmp_pair(p, q, quick)
        except Exception as e:
            add(f"  !! {p.name[:44]}  讀取失敗: {type(e).__name__}: {e}")
            err += 1
            continue
        if probs:
            diff += 1
            add(f"  !! {p.name[:40]}  <->  {q.name[:40]}")
            for x in probs[:MAX_SHOW]:
                add(f"       {x}")
        else:
            same += 1
    bar.close()
    add("")
    add(f"  一致 {same} 組   有差異 {diff} 組   讀不到 {err} 組")
    if _NEAR["multi"]:
        add("")
        add(f"  註: 有 {_NEAR['multi']:,} 個格子是把多值欄位的元素排序之後才")
        add("      對上的。GROUP_CONCAT 沒有 ORDER BY 時元素順序未定義,")
        add("      兩個引擎排出來不一樣, 但裝的是同一組值。")
        add("      底稿要寫明這一項用的是無序比對 (--loose)。")
    if _NEAR["n"]:
        add("")
        add(f"  註: 有 {_NEAR['n']:,} 個數字是取 {_CMP_SIG} 位有效位數才對上的,")
        add(f"      最大相對差 {_NEAR['max']:.2e}。")
        add("      這是 SUM() 的浮點累加順序造成的 (sqlite 逐列, DuckDB 平行),")
        add("      不是邏輯差異。底稿要寫明容差, 不要當成完全逐字一致。")
        if _NEAR["max"] > 1e-9:
            add("      !! 超過 1e-9, 這不像單純的累加順序, 要查清楚。")
    if pairs and diff == 0 and err == 0:
        add("  配對到的全部一致 —— 轉 parquet 與改路徑沒有影響內容。")




# ------------------------------------------------------- --check
def mode_check(root: Path) -> None:
    """開跑之前的體檢。

    三支腳本會讀「自己所在資料夾的 data/」、寫「output/」。放錯地方、混到
    不該進來的檔、lookup 少一個, 都會讓整輪跑完才發現結果不對。這些全部
    是秒級就查得出來的, 跑之前先過一遍。
    """
    add("")
    add("=" * 72)
    add(f"開跑前體檢   {root}")
    add("=" * 72)
    bad = 0

    def ok(msg):
        add(f"  OK   {msg}")

    def no(msg, hint=""):
        nonlocal bad
        bad += 1
        add(f"  !!   {msg}")
        if hint:
            add(f"       {hint}")

    # 1. 三支腳本在不在
    found = {}
    for hint, what in ((COMBINE_HINT, "合併/重複偵測"),
                       (ARA_HINT, "月度彙總"), (ADJ_HINT, "案件彙總")):
        ps = _patch_find(root, hint)
        found[what] = ps
        if ps:
            ok(f"{what}腳本: {ps[0].name[:50]}")
        else:
            no(f"找不到{what}腳本 (檔名要含 {hint!r})")

    # 2. 改過了沒
    for what, ps in found.items():
        for q in ps:
            t = q.read_text(encoding="utf-8", errors="replace")
            has_pq = "PARQUET_EXTENSIONS" in t or "iter_batches" in t
            has_path = '/ "data"' in t or '/ "output"' in t
            if what == "案件彙總":
                has_pq = True      # 這支不讀原始資料, 不需要 parquet 支援
            if has_pq and has_path:
                ok(f"{q.name[:46]} 已改好 (parquet + 路徑)")
            else:
                miss = []
                if not has_pq:
                    miss.append("parquet")
                if not has_path:
                    miss.append("data/output 路徑")
                no(f"{q.name[:46]} 還沒改: 缺 {', '.join(miss)}",
                   "先跑  python csvtools.py --patch --paths")

    # 3. data/ 裡面有什麼
    d = root / DATA_DIR
    if not d.is_dir():
        no(f"沒有 {DATA_DIR}/ 資料夾")
    else:
        pq_ = list(d.glob("*.parquet"))
        csv_ = list(d.glob("*.csv"))
        xl_ = [x for x in d.glob("*.xls*") if not x.name.startswith("~$")]
        sub = [x for x in d.iterdir() if x.is_dir()]
        ok(f"{DATA_DIR}/  parquet {len(pq_)}  csv {len(csv_)}  excel {len(xl_)}")
        if not pq_:
            no(f"{DATA_DIR}/ 裡沒有 parquet")
        lower = {x.name.lower() for x in csv_}
        for want in ("fx rate.csv", "station region mapping.csv"):
            if want in lower:
                ok(f"{DATA_DIR}/{want} 在")
            else:
                no(f"{DATA_DIR}/ 缺 {want}",
                   "合併腳本找不到它會直接停下來" if "fx" in want else "")
        extra = [x for x in csv_ if x.name.lower() not in
                 ("fx rate.csv", "station region mapping.csv")]
        if extra:
            no(f"{DATA_DIR}/ 還有 {len(extra)} 個 csv 會被當成來源資料讀:",
               ", ".join(x.name[:40] for x in extra[:6]))
        if xl_:
            no(f"{DATA_DIR}/ 有 {len(xl_)} 個 Excel 也會被讀進來:",
               ", ".join(x.name[:40] for x in xl_[:6]))
        if sub:
            add(f"  --   {DATA_DIR}/ 底下還有 {len(sub)} 個子資料夾"
                " (不會被讀, 腳本只掃第一層)")
        # parquet 是壓縮的, 檔案大小不等於處理時的資料量。sqlite 存的是
        # 未壓縮的文字, 所以要看 parquet metadata 裡的「解壓後大小」才估得準。
        # metadata 不用讀資料, 秒級。
        if pq_:
            try:
                import pyarrow as pa_
                import pyarrow.parquet as pqm
                # 列數從 metadata 拿, 秒級。每列多少位元組要實際量 ——
                # metadata 裡的 total_byte_size 是「編碼後」的大小, 字典編碼
                # 之後遠小於真正的文字量, 拿來估 sqlite 會低估一個數量級。
                nrow = 0
                for q in pq_:
                    nrow += pqm.ParquetFile(q).metadata.num_rows
                # 量「實際文字量」, 不要用 arrow 的 nbytes —— 那裡面每個值
                # 有 4 bytes 的 offset, 幾十個欄位就是每列兩百多 bytes 的純
                # 開銷, 拿來估 sqlite 會高估將近一倍。
                import pyarrow.compute as pc_
                per, ncol, n_s = 0.0, 0, 0
                for q in pq_[:: max(1, len(pq_) // 5)][:5]:
                    b = next(pqm.ParquetFile(q).iter_batches(batch_size=20000))
                    t = pa_.Table.from_batches([b])
                    if not t.num_rows:
                        continue
                    txt = 0
                    for c in t.columns:
                        try:
                            txt += pc_.sum(pc_.binary_length(c)).as_py() or 0
                        except Exception:
                            txt += c.nbytes
                    per += txt / t.num_rows
                    ncol = max(ncol, t.num_columns)
                    n_s += 1
                per = per / n_s if n_s else 0
                # sqlite 每個欄位值再加 1~2 bytes 的型別標頭
                row_b = per + ncol * 1.5
                tbl = int(nrow * row_b)
                idx = int(nrow * 30 * 3)      # combined 上那三個索引
                add(f"  --   {DATA_DIR}/ 共 {nrow:,} 列   壓縮後 "
                    f"{mb(sum(q.stat().st_size for q in pq_))}   "
                    f"純文字約 {mb(int(nrow * per))}"
                    f"  ({per:.0f} bytes/列 × {ncol} 欄, 取樣 {n_s} 檔)")
                add(f"       灌進 sqlite 粗估: 資料 {mb(tbl)} + 索引 {mb(idx)}"
                    f" = {mb(tbl + idx)}")
                add(f"       (合併腳本另有暫存表, 實際峰值會再高一些)")
            except Exception as e:
                add(f"  --   算不出資料量: {type(e).__name__}: {e}")

        # 檔數對不對, 光看數字看不出來 —— 攤成月份才知道有沒有缺
        spans = [parse_span(x.name)[0] for x in pq_]
        months = sorted({m for sp in spans if sp for m in sp})
        nobody = [x.name for x in pq_ if parse_span(x.name)[0] is None]
        if months:
            span_all = []
            y, mo = months[0]
            while (y, mo) <= months[-1]:
                span_all.append((y, mo))
                y, mo = (y + 1, 1) if mo == 12 else (y, mo + 1)
            miss = [x for x in span_all if x not in set(months)]
            add(f"  {'OK' if not miss else '!!'}   {DATA_DIR}/ 涵蓋 "
                f"{months[0][0]}-{months[0][1]:02d} ~ "
                f"{months[-1][0]}-{months[-1][1]:02d}, "
                f"{len(months)}/{len(span_all)} 個月")
            if miss:
                bad += 1
                add("       缺: " + ", ".join(f"{a}-{b:02d}" for a, b in miss[:18])
                    + (" …" if len(miss) > 18 else ""))
            add("       (這是全部檔合起來看的。個別系列有沒有洞要看"
                " --coverage)")
        if nobody:
            add(f"  --   {len(nobody)} 個檔的檔名看不出期間: "
                + ", ".join(x[:36] for x in nobody[:4]))

    # 4. output/ 與對方產出
    o = root / OUT_DIR
    if o.is_dir():
        n = len([x for x in o.iterdir() if x.is_file()])
        add(f"  --   {OUT_DIR}/ 已存在, 裡面有 {n} 個檔"
            + ("  (重跑會覆蓋)" if n else ""))
    else:
        add(f"  --   {OUT_DIR}/ 還沒有, 腳本會自己建")
    m = root / MISC
    if m.is_dir():
        n = len([x for x in m.iterdir()
                 if x.suffix.lower() in (".xlsx", ".csv")])
        ok(f"{MISC}/ 有 {n} 個可比對的檔 (--cmp 用)")
    else:
        add(f"  --   沒有 {MISC}/, 之後 --cmp 要用 --theirs 指定")

    add("")
    add("  可以跑了。" if not bad else f"  有 {bad} 項要先處理。")




# ------------------------------------------------------- --scope
def mode_scope(root: Path, theirs: str) -> None:
    """對方那幾份產出各自涵蓋哪些來源檔。

    磁碟放不下一次跑 260 個檔, 而對方本來就有兩份 Output_analysis, 代表
    他們是分範圍跑的。產出裡有 Source File 欄, 把它的相異值撈出來, 就知道
    每一輪該餵哪些檔 —— 照著切, 既放得下, 比對也才對得上。
    """
    from openpyxl import load_workbook
    d = root / theirs
    if not d.is_dir():
        add(f"  !! 找不到 {d}")
        return
    add("")
    add("=" * 72)
    add(f"對方產出的來源範圍   掃 {theirs}/")
    add("=" * 72)
    fs = [p for p in sorted(d.iterdir())
          if p.suffix.lower() in (".xlsx", ".xlsm") and not p.name.startswith("~$")]
    for f in tqdm(fs, desc="  讀取", unit="檔", leave=False):
        try:
            wb = load_workbook(f, read_only=True, data_only=True)
        except Exception as e:
            add(f"  !! {f.name[:50]}  讀不到: {type(e).__name__}")
            continue
        hits: dict = {}
        try:
            for ws in wb.worksheets:
                it = ws.iter_rows(values_only=True)
                hdr = ["" if v is None else str(v).strip() for v in next(it, ())]
                idx = next((i for i, c in enumerate(hdr)
                            if "source" in c.lower() and "file" in c.lower()), None)
                if idx is None:
                    continue
                vals = set()
                for row in it:
                    if row is None or idx >= len(row):
                        continue
                    v = row[idx]
                    if v not in (None, ""):
                        vals.add(str(v).strip())
                if vals:
                    hits[ws.title] = vals
        finally:
            wb.close()
        if not hits:
            add(f"  -- {f.name[:56]}  沒有 Source File 欄")
            continue
        allv = set().union(*hits.values())
        add("")
        add(f"  {f.name}")
        add(f"    {len(allv)} 個來源檔 (出現在 {len(hits)} 張工作表)")
        ser: Counter = Counter(series_of(x) for x in allv)
        for sn, n in ser.most_common():
            add(f"      {n:>4} 檔   [{sn}]")
        for x in sorted(allv)[:6]:
            add(f"        {x[:64]}")
        if len(allv) > 6:
            add(f"        ... 另外 {len(allv) - 6} 個")



# ---------------------------------------------------------------- --fc
# 兩邊各自從系統拉出來的檔要證明「內容一樣」, 但檔名規則、檔案格式、切檔
# 方式三樣都不同, 沒辦法比位元組。這裡的做法是:
#   1. 用檔名推出每個檔涵蓋哪些月份 + 顆粒度 (parse_span)
#   2. (系列, 期間, 顆粒度) 一樣的檔先併成一個「邏輯單位」—— 對方把一個月
#      拆成 _01/_02 兩份, 那兩份合起來才等於我方的一個檔
#   3. 期間完全一樣而且兩邊各只有一個 -> 直接配對, 順便記下「ita 的這個系列
#      對應對方的哪個系列」
#   4. 剩下的 (我方單月 vs 對方整季/整年) 用學到的系列對應 + 期間包含關係配
#   5. 比內容: 把整列接成一個字串取雜湊, 丟進 Counter 相減。順序不影響。
FC = "filechecks"
FC_DIFF = "fc_diffs.csv"      # 對不上的列寫這裡 (有票號, 內部用)
FC_DIFF_CAP = 5000            # 最多寫幾列, 免得炸出一個巨檔
FC_DIAG_CAP = 20_000          # 對不上的列超過這個數就不做逐欄暴力診斷
FC_KEY_CAP = 300_000          # 再多就連湊對都放棄, 只留 fc_diffs.csv
FC_SHOW = 10                  # 對不上的列少於這個數就直接印出來 (遮罩過)


def _fc_sort(x):
    return sorted(x[0]["span"]), FC_COARSE.get(x[0]["gran"], 9)

# 顆粒度由細到粗。用來判斷「對方那一份涵蓋得比較廣」。
FC_COARSE = {"上半月": 0, "下半月": 0, "部分月": 0, "整月": 1, "月範圍": 2,
             "1Q": 2, "2Q": 2, "3Q": 2, "4Q": 2, "1H": 3, "2H": 3, "整年": 4}


def fc_batches(p: Path, bad: list | None = None):
    """串流讀一個檔, 逐批回 pyarrow Table (全部欄位都是文字)。

    bad 給一個 list 的話, 讀不動的列會跳過並記進去, 而不是整個檔放棄。
    一個檔裡有幾列壞掉就整份不比, 等於用「讀不到」蓋掉「內容不一樣」。
    """
    import pyarrow as pa
    if p.suffix.lower() == ".parquet":
        import pyarrow.parquet as pq
        for b in pq.ParquetFile(p).iter_batches(batch_size=100_000):
            yield pa.Table.from_batches([b])
    else:
        import pyarrow.csv as pacsv
        # 一律當文字讀。讓 pyarrow 推型別的話, 票號的前導零會掉, 兩邊就算
        # 內容一樣也會比出差異 —— 而且是靜靜地錯。
        co = pacsv.ConvertOptions(
            column_types={c: pa.string() for c in csv_header(p)})
        # 欄位值裡可能夾著換行 (備註欄那種)。預設會把它當成新的一列, 於是
        # 報「Expected 54 columns, got 19」然後整個檔讀不動。
        kw = {"newlines_in_values": True}
        if bad is not None:
            def keep_going(row):
                bad.append((row.number, row.actual_columns,
                            (row.text or "")[:70]))
                return "skip"
            kw["invalid_row_handler"] = keep_going
        try:
            po = pacsv.ParseOptions(**kw)
        except TypeError:                      # 舊版 pyarrow 沒有這個參數
            po = pacsv.ParseOptions(newlines_in_values=True)
        for b in pacsv.open_csv(
                p, read_options=pacsv.ReadOptions(block_size=1 << 26),
                parse_options=po, convert_options=co):
            yield pa.Table.from_batches([b])


def fc_cols(u: dict) -> list[str]:
    p = u["files"][0]
    if p.suffix.lower() == ".parquet":
        import pyarrow.parquet as pq
        return list(pq.ParquetFile(p).schema_arrow.names)
    return csv_header(p)


def fc_units(base: Path) -> list[dict]:
    out: dict = {}
    for p in sorted(base.iterdir()):
        if p.is_dir() or p.suffix.lower() not in (".csv", ".parquet"):
            continue
        span, gran = parse_span(p.name)
        key = (series_of(p.name), frozenset(span or ()), gran or "?")
        out.setdefault(key, {"series": key[0], "span": key[1],
                             "gran": key[2], "files": []})["files"].append(p)
    for u in out.values():
        u["bytes"] = sum(f.stat().st_size for f in u["files"])
    return list(out.values())


def fc_label(u: dict) -> str:
    s = sorted(u["span"])
    if not s:
        return "?"
    if len(s) == 1:
        return f"{s[0][0]}-{s[0][1]:02d} {u['gran']}"
    return f"{s[0][0]}-{s[0][1]:02d}..{s[-1][0]}-{s[-1][1]:02d} {u['gran']}"


def fc_names(u: dict) -> str:
    if len(u["files"]) == 1:
        return u["files"][0].name
    return f"{u['files'][0].name}  (+{len(u['files']) - 1} 份)"


def fc_pair(ita: list, cli: list):
    """回 (配對清單, ita 沒對手, client 沒對手, 系列對應)。

    只看期間會出事: 兩邊各有一個系列剛好都是「2014 年 11 月的整月檔」, 但
    但兩者其實是不同的抽取, 期間一樣不代表是同一批資料。所以
    兩邊有同名系列時直接鎖定, 不准跨系列配。
    """
    cli_series = {x["series"] for x in cli}
    vote: dict = {u["series"]: {u["series"]}
                  for u in ita if u["series"] in cli_series}

    bucket: dict = {}
    for x in cli:
        bucket.setdefault((x["span"], x["gran"]), []).append(x)
    pairs, left, used = [], [], set()

    # 第一輪 —— 期間與顆粒度完全一樣, 而且兩邊各只有一個, 沒有別的可能
    for u in ita:
        allow = vote.get(u["series"])
        c = [x for x in bucket.get((u["span"], u["gran"]), [])
             if id(x) not in used and (not allow or x["series"] in allow)]
        if len(c) == 1:
            used.add(id(c[0]))
            pairs.append((u, c[0], "同期"))
            vote.setdefault(u["series"], set()).add(c[0]["series"])
        else:
            left.append(u)

    # 第二輪 —— 我方是單月, 對方沒有單獨拆出來 (整季/整年)。用學到的系列
    # 對應把候選縮小, 再取期間最窄的那個。對方那個大檔可能同時是好幾個小檔
    # 的對手, 所以這裡不標 used。
    #
    # 「涵蓋範圍比較大」必須真的比較大: 月份集合是真子集, 或月份一樣但顆粒
    # 度更粗 (我方上半月 vs 對方整月)。少了這一關, 上半月會被配到下半月身
    # 上 —— 月份集合一模一樣, 光看包含關係分不出來。
    for u in list(left):
        allow = vote.get(u["series"])
        cands = [x for x in cli
                 if u["span"]
                 and (u["span"] < x["span"]
                      or (u["span"] == x["span"]
                          and FC_COARSE.get(x["gran"], 9) > FC_COARSE.get(u["gran"], 9)))
                 and (not allow or x["series"] in allow)]
        if cands:
            best = sorted(cands, key=lambda x: (len(x["span"]), x["series"]))[0]
            pairs.append((u, best, "對方沒單獨拆"))
            left.remove(u)

    matched = {id(p[1]) for p in pairs}
    return (pairs, left, [x for x in cli if id(x) not in matched], vote)


def fc_rowvals(u: dict, cols: list[str], want: set, cap: int,
               only_ym=None, loose=None, deg=None) -> list:
    """撈出對不上的列。only_ym = (欄index, (年,月)) 時只收那個月的。

    對方那份涵蓋整年時, 「對不上」的列有十一個月是本來就該多的。不篩掉的話
    診斷會被這堆無關的列淹掉, 而且記憶體也吃不消。
    """
    out = []
    for f in u["files"]:
        for t in fc_batches(f):
            for s_ in fc_joined(t, cols, loose, deg):
                if hash(s_) not in want:
                    continue
                r = s_.split("\x1f")
                if only_ym and _ym(r[only_ym[0]]) != only_ym[1]:
                    continue
                out.append(r)
                if len(out) >= cap:
                    return out
    return out


def fc_explain(cols: list[str], ra: list, rb: list):
    """逐欄診斷 —— 每次拿掉一欄重算, 看能對上多少列。

    不用猜哪一欄是主鍵, 也不怕差異橫跨好幾欄: 能把對不上的列解釋掉的那一
    欄, 就是兩邊寫法不同的那一欄。順便撈一組實際的值回來 —— 只知道「是這
    一欄」還不夠, 要看到兩邊各是什麼才判斷得出是格式問題還是資料問題。
    """
    hits = []
    for i, c in enumerate(cols):
        a: dict = {}
        b: dict = {}
        for r in ra:
            a.setdefault(tuple(r[:i] + r[i + 1:]), []).append(r[i])
        for r in rb:
            b.setdefault(tuple(r[:i] + r[i + 1:]), []).append(r[i])
        common = a.keys() & b.keys()
        if not common:
            continue
        n = sum(min(len(a[k]), len(b[k])) for k in common)
        k0 = next(iter(common))
        hits.append((n, i, c, a[k0][0], b[k0][0]))
    hits.sort(key=lambda x: -x[0])
    return hits


def fc_keycols(cols: list[str]) -> list[int]:
    """挑一組當「同一列」的鍵。找不到就回空的。"""
    idx = []
    for want in (("tkt", "ticket"), ("coupon",)):
        for i, c in enumerate(cols):
            k = c.lower()
            if any(w in k for w in want) and ("num" in k or "no" in k
                                              or "coupon" in k):
                idx.append(i)
                break
    return idx


def fc_bykey(cols: list[str], ra: list, rb: list, keyidx: list):
    """列太多不能逐欄暴力時, 用票號+航段湊對, 再看是哪幾欄不同。"""
    b: dict = {}
    for r in rb:
        b.setdefault(tuple(r[i] for i in keyidx), []).append(r)
    percol: Counter = Counter()
    sample: dict = {}
    paired = nokey = 0
    for r in ra:
        cand = b.get(tuple(r[i] for i in keyidx))
        if not cand:
            nokey += 1
            continue
        best = min(cand, key=lambda x: sum(1 for i in range(len(cols))
                                           if x[i] != r[i]))
        d = [i for i in range(len(cols)) if best[i] != r[i]]
        if d:
            paired += 1
            for i in d:
                percol[cols[i]] += 1
                sample.setdefault(cols[i], (r[i], best[i]))
    return paired, nokey, percol, sample


def _ym(v):
    if not v:
        return None
    m = re.match(r"(20\d{2})[-/]?(\d{2})", str(v))
    if m and 1 <= int(m.group(2)) <= 12:
        return int(m.group(1)), int(m.group(2))
    return None


def fc_pick_period(cols: list[str]):
    low = {c.lower(): c for c in cols}
    for k in low:
        if "yearmonth" in k or "year_month" in k or "yr_month" in k:
            return low[k]
    for k in low:
        if "date" in k and any(w in k for w in ("flight", "uplift", "dep")):
            return low[k]
    return None


def fc_periods(u: dict, cols: list[str]):
    """對方那一份實際涵蓋哪些月份。回 (用了哪一欄, {(年,月): 列數})。"""
    pick = fc_pick_period(cols)
    if not pick:
        return None, None
    cnt: Counter = Counter()
    for f in u["files"]:
        for t in fc_batches(f):
            if pick not in t.column_names:
                continue
            for v in t.column(pick).to_pylist():
                ym = _ym(v)
                if ym:
                    cnt[ym] += 1
    return pick, cnt


def fc_joined(tbl, cols, loose=None, deg=None) -> list[str]:
    """整列接成一個字串。空值一律當空字串 —— parquet 的 null 與 CSV 的空
    欄位是同一件事, 不統一的話兩邊會無謂地比不中。

    loose 裡的欄位會先把 "|" 串起來的多值拆開排序再接回去: 那種欄位兩邊
    可能裝著同樣的值卻排不同順序, 逐字比會全部比不中, 但那不是資料不同。
    """
    import pyarrow as pa
    import pyarrow.compute as pc
    arrs = []
    for c in cols:
        a = tbl.column(c)
        if deg and c in deg:
            fn = deg[c][1]
            a = pa.array([fn(v) if v is not None else "" for v in a.to_pylist()],
                         type=pa.string())
        if loose and c in loose:
            a = pa.array(["" if v is None else "|".join(sorted(v.split("|")))
                          for v in a.to_pylist()], type=pa.string())
        arrs.append(a if a.type == pa.string() else a.cast(pa.string()))
    return pc.binary_join_element_wise(
        *arrs, "\x1f", null_handling="replace",
        null_replacement="").to_pylist()


def fc_hash(tbl, cols, loose=None, deg=None):
    return map(hash, fc_joined(tbl, cols, loose, deg))


def fc_multi(u: dict, cols: list[str]) -> set:
    """哪些欄位裝的是 "|" 串起來的多值。只看第一批就夠判斷。"""
    t = next(fc_batches(u["files"][0]), None)
    if t is None:
        return set()
    out = set()
    for c in cols:
        if c not in t.column_names:
            continue
        for v in t.column(c)[:3000].to_pylist():
            if v and "|" in v:
                out.add(c)
                break
    return out


def fc_same_set(x: str, y: str) -> bool:
    """兩個值是不是同一組多值, 只是順序不同。"""
    return ("|" in x or "|" in y) and sorted(x.split("|")) == sorted(y.split("|"))


def fc_colcompare(cols: list[str], ra: list, rb: list, picks: list):
    """低重疊的那幾欄, 兩邊各長什麼樣 —— 字元形狀 + 值域。

    「格式不同」跟「根本是不同的資料」在重疊率上長得一樣 (都是 0%), 但
    看值域就分得出來: 形狀不同是格式問題; 形狀一樣而值域錯開, 就是兩批
    不同期間/不同來源的資料。
    """
    out = []
    for c in picks:
        i = cols.index(c)
        va = [r[i] for r in ra[:20000] if r[i]]
        vb = [r[i] for r in rb[:20000] if r[i]]
        sa = Counter(fc_shape(v) for v in va).most_common(1)
        sb = Counter(fc_shape(v) for v in vb).most_common(1)
        na_, xa_ = _rng(va)
        nb_, xb_ = _rng(vb)
        out.append((c, sa[0][0] if sa else "(全空)", na_, xa_,
                    sb[0][0] if sb else "(全空)", nb_, xb_))
    return out


def _rng(v: list):
    """值域。金額那種欄位照字串排會排出 "100.00 ~ 99.00" 這種看了會誤會的
    結果, 全部是數字就照數字排。"""
    if not v:
        return "", ""
    try:
        f = [float(x) for x in v]
        return v[f.index(min(f))], v[f.index(max(f))]
    except ValueError:
        return min(v), max(v)


# ---- 降階比對 ----------------------------------------------------------
# 一邊的檔被 Excel 改過 (13 位數字變科學記號、日期換格式、秒被捨掉) 時, 壞
# 掉的值還原不回來 —— 科學記號只剩六位有效數字。但可以反過來: 把完好的那
# 一邊也降到同樣的精度再比。比中了只能說「在受損檔還保留的精度上沒有差
# 異」, 不等於逐字一致, 這個限制必須寫在結論裡。
def _deg_num(v):
    if not v or "." not in v:
        return v
    try:
        float(v)
    except ValueError:
        return v
    return v.rstrip("0").rstrip(".")


def _deg_date(v):
    m = re.match(r"(\d{4})-(\d{2})-(\d{2})$", v or "")
    return f"{int(m.group(2))}/{int(m.group(3))}/{m.group(1)}" if m else v


def _deg_dt(v):
    m = re.match(r"(\d{4})-(\d{2})-(\d{2})[ T](\d{1,2}):(\d{2})", v or "")
    return (f"{int(m.group(2))}/{int(m.group(3))}/{m.group(1)} "
            f"{int(m.group(4))}:{m.group(5)}") if m else v


def _deg_tkt(cut):
    """Excel 的 General 格式: 位數少的照原樣顯示, 超過欄寬就轉科學記號。
    臨界點看欄寬, 從資料本身推不出來, 所以幾個都試, 由形狀比對決定哪個對。
    """
    def f(v):
        if not v or not v.isdigit():
            return v
        n = int(v)
        t = str(n)
        return t if len(t) <= cut else f"{n:.5E}"
    return f


def _canon_num(v):
    """兩邊都轉成數值再比 —— 只把「同一個數字的不同寫法」視為相同
    (0.00 = 0 = 0.0), 不同的數字還是不同。不失真。"""
    if not v:
        return v
    try:
        f = float(v)
    except ValueError:
        return v
    return repr(int(f)) if f == int(f) else repr(f)


def _canon_sig6(v):
    """兩邊都降到 6 位有效數字。

    一邊已經是科學記號時, 尾數是真的沒了 —— 單向轉換猜不出 Excel 當初
    的臨界點 (那取決於當時的欄寬)。所以兩側一起降到同一個精度, 不用猜。
    代價是這一欄變成**會把不同的值視為相同**, 結論只能講到這個精度為止。
    """
    if not v:
        return v
    try:
        return f"{float(v):.5E}"
    except ValueError:
        return v


def _numfrac(vals) -> float:
    n = ok = 0
    for v in vals:
        if not v:
            continue
        n += 1
        try:
            float(v)
            ok += 1
        except ValueError:
            pass
    return ok / n if n else 0.0


DEGRADES = ([("去掉尾隨的零 (0.00 -> 0)", _deg_num),
             ("ISO -> M/D/YYYY H:MM (捨去秒)", _deg_dt),
             ("ISO -> M/D/YYYY", _deg_date)]
            + [(f"去前導零, {c + 1} 位以上轉科學記號", _deg_tkt(c))
               for c in range(9, 14)])


def fc_degrade_plan(pa_: Path, pb_: Path, cols: list[str]) -> dict:
    """哪些欄要把 b 側降階, 用哪一種。

    只有「降階之後 b 側的字元形狀跟 a 側完全一樣」才採用 —— 猜錯就不套,
    不會為了湊成一樣而硬改資料。
    """
    ta, tb = fc_head(pa_), fc_head(pb_)
    if ta is None or tb is None:
        return {}
    plan = {}
    for c in cols:
        if c not in ta.column_names or c not in tb.column_names:
            continue
        va = [v for v in ta.column(c)[:3000].to_pylist() if v]
        vb = [v for v in tb.column(c)[:3000].to_pylist() if v]
        if not va or not vb:
            continue
        want = {fc_shape(v) for v in va}
        if want == {fc_shape(v) for v in vb}:
            continue
        # 某一邊已經是科學記號 = 尾數真的沒了。這種情況單向重建不可能準確
        # (臨界點取決於當初的欄寬, 猜中一個「看起來合理」的值反而更危險 ——
        # 形狀對得上但有一部分列還是錯的)。直接兩側一起降到同一個精度。
        if any(fc_smell(x) for x in want | {fc_shape(v) for v in vb}):
            if _numfrac(va) >= 0.95 and _numfrac(vb) >= 0.95:
                plan[c] = ("兩邊都降到 6 位有效數字 (科學記號已無法還原)",
                           _canon_sig6, True)
                continue
        hit = False
        for label, fn in DEGRADES:
            got = {fc_shape(fn(v)) for v in vb}
            # 要求「a 側出現過的形狀, 轉換後的 b 側都產得出來」而不是完全
            # 相等 —— b 側那個檔常常涵蓋更多月份, 十二月的日期比九月多一
            # 位數, 硬要求相等會把正確的轉換也擋掉。真正的證明是後面的逐
            # 列比對, 這裡只是挑轉換用的。
            if want <= got and got != {fc_shape(v) for v in vb}:
                plan[c] = (label, fn, False)
                hit = True
                break
        if hit:
            continue
        # 單向轉換套不上, 但兩邊都是數字的話, 可以兩側一起正規化。方向對
        # 稱, 不會偏向任何一邊, 也不必猜對方當初做了什麼。
        if _numfrac(va) >= 0.95 and _numfrac(vb) >= 0.95:
            if any(fc_smell(x) for x in want | {fc_shape(v) for v in vb}):
                plan[c] = ("兩邊都降到 6 位有效數字 (科學記號已無法還原)",
                           _canon_sig6, True)
            else:
                plan[c] = ("兩邊都轉成數值再比 (0.00 = 0)", _canon_num, True)
    return plan


def fc_overlap(cols: list[str], ra: list, rb: list):
    """每一欄各自的值重疊多少。整列湊不起來的時候, 這個看得出是哪幾欄不同
    ——全部欄位都不重疊代表根本是兩批資料, 只有一兩欄不重疊就是那幾欄的
    問題。"""
    out = []
    for i, c in enumerate(cols):
        a = Counter(r[i] for r in ra)
        b = Counter(r[i] for r in rb)
        out.append((sum((a & b).values()) / max(1, len(ra)), c))
    out.sort()
    return out


def fc_examples(u: dict, cols: list[str], want: set, cap: int,
                loose=None, deg=None) -> list:
    """把雜湊落在 want 裡的列原文撈回來, 給人去 raw file 追。"""
    out = []
    for f in u["files"]:
        for t in fc_batches(f):
            for s_ in fc_joined(t, cols, loose, deg):
                if hash(s_) in want:
                    out.append((f.name, dict(zip(cols, s_.split("\x1f")))))
                    if len(out) >= cap:
                        return out
    return out


def fc_cname(c: str) -> str:
    """欄名可能是空白的 (匯出時多一個逗號就會這樣)。直接印會變成一個洞,
    看不出在講哪一欄。"""
    return repr(c) if c.strip() else "(名稱空白)"


def fc_short(name: str, prefix: str) -> str:
    n = name[len(prefix):] if prefix and name.startswith(prefix) else name
    return n[:42]


def fc_count(u: dict, bar=None) -> int:
    import pyarrow.parquet as pq
    n = 0
    for p in u["files"]:
        if p.suffix.lower() == ".parquet":
            n += pq.ParquetFile(p).metadata.num_rows
            if bar:
                bar.update(p.stat().st_size)
        else:
            for t in fc_batches(p):
                n += t.num_rows
            if bar:
                bar.update(p.stat().st_size)
    return n


def fc_shape(v) -> str:
    if v is None:
        return "(空)"
    t = re.sub(r"\d", "9", str(v))
    return re.sub(r"[A-Za-z]", "A", t)[:24]


def fc_head(p: Path):
    """只讀開頭一小塊。看寫法用的, 不必把 64MB 的區塊整個讀進來。"""
    import pyarrow as pa
    if p.suffix.lower() == ".parquet":
        import pyarrow.parquet as pq
        for b in pq.ParquetFile(p).iter_batches(batch_size=4000):
            return pa.Table.from_batches([b])
        return None
    import pyarrow.csv as pacsv
    co = pacsv.ConvertOptions(
        column_types={c: pa.string() for c in csv_header(p)})
    for b in pacsv.open_csv(
            p, read_options=pacsv.ReadOptions(block_size=1 << 20),
            convert_options=co):
        return pa.Table.from_batches([b])
    return None


def fc_smell(shape: str):
    """看形狀就知道這一欄被什麼東西改過。"""
    if re.fullmatch(r"9\.9+A\+99", shape):
        return ("科學記號 —— 13 位數字用 Excel 開過就會變成這樣, "
                "而且尾數回不來了")
    return None


def fc_shapes(p: Path, cols: list[str], n: int = 4000) -> dict:
    """每欄回 [(形狀, 次數), ...] 由多到少。只讀開頭, 夠看寫法了。"""
    t = fc_head(p)
    out: dict = {}
    if t is None:
        return out
    for c in cols:
        if c not in t.column_names:
            continue
        out[c] = Counter(fc_shape(v) for v in t.column(c)[:n].to_pylist())
    return out


def fc_shape_diff(sa: dict, sb: dict, cols: list[str]) -> list:
    """只挑真的有問題的欄。回 [(欄, ita主要寫法, client主要寫法, 為什麼)]。

    兩邊是從各自檔案的開頭取樣, 而且兩個檔的列序本來就不一樣 —— 某一欄在
    這幾千列裡剛好多半是空的、另一邊剛好多半有值, 是取樣的常態, 不是資料
    差異。所以只認兩種情況:

      1. 某一邊出現看得出被改過的寫法 (例如科學記號), 不管佔比多少
      2. 兩邊各自都有一種壓倒性的寫法 (九成以上), 而那兩種不一樣

    其餘一律不報。寧可漏掉邊緣情況, 也不要每一組都亮紅燈 —— 紅燈滿天飛
    跟沒有紅燈是一樣的。
    """
    out = []
    for c in cols:
        a, b = sa.get(c), sb.get(c)
        if not a or not b:
            continue
        na, nb = sum(a.values()), sum(b.values())
        if not na or not nb:
            continue
        ta, ka = a.most_common(1)[0]
        tb, kb = b.most_common(1)[0]
        why = None
        for side, cnt, tot in (("ita", a, na), ("client", b, nb)):
            for sh, k in cnt.items():
                if k / tot >= 0.01 and fc_smell(sh):
                    why = f"{side} 這一欄是{fc_smell(sh)}"
                    break
            if why:
                break
        if (not why and ta and tb and ta != tb
                and ka / na >= 0.9 and kb / nb >= 0.9):
            why = "兩邊各自都是單一寫法, 但不是同一種"
        if why:
            out.append((c, ta or "(空)", tb or "(空)", why))
    return out


_MASKMAP: dict = {}


def fc_mask(c: str, v: str, raw: bool = True) -> str:
    """票號換成穩定編號 —— 只有 --mask 才會用到。

    預設是原樣顯示: 這幾列是拿來 debug 的, 數量也不多, 遮了反而查不下去。
    要把輸出貼到外面才加 --mask, 同一個票號會對到同一個 #n, 一樣就是一
    樣、不一樣就是不一樣, 但號碼本身不外流。
    """
    k = c.lower()
    if not v or raw:
        return v
    if "tkt" in k or "ticket" in k or re.fullmatch(r"\d{9,}", v):
        if v not in _MASKMAP:
            _MASKMAP[v] = f"#{len(_MASKMAP) + 1}"
        return _MASKMAP[v]
    return v


def fc_findkey(u: dict, cols: list[str], keyidx: list, keys: set,
               cap: int, only_ym=None, loose=None, deg=None) -> list:
    """在對方那一份裡撈出「鍵相同」的列 —— 其他欄位一不一樣都撈。

    雜湊比對只會說「這一列對方沒有」, 但那有兩種情況: 對方連這張票都沒
    有, 還是有這張票但某一欄的值不同。差別很大, 得用鍵去找才知道。
    """
    out = []
    for f in u["files"]:
        for t in fc_batches(f):
            for s_ in fc_joined(t, cols, loose, deg):
                r = s_.split("\x1f")
                if only_ym and _ym(r[only_ym[0]]) != only_ym[1]:
                    continue
                if tuple(r[i] for i in keyidx) in keys:
                    out.append(r)
                    if len(out) >= cap:
                        return out
    return out


# ------------------------------------------------- --fc --flat
# ita 這次是把 31 個檔平鋪在 filechecks/ 底下, 沒有 ita/ client/ 兩層, 檔名
# 也換了新的寫法 ("2026 version", "KA with 160 stock")。靠檔名配對會出事 ——
# 上次 043 月檔就是這樣配錯, 0% 重疊白跑一輪。
#
# 所以改用內容配對: 每個檔算出「票號+票聯」這組鍵的集合, 兩邊取交集, 重疊率
# 最高的才算一對。檔名只拿來顯示, 不參與判斷。
def fcf_files(d: Path, exts: tuple, deep: bool) -> list:
    it = d.rglob("*") if deep else d.iterdir()
    out = []
    for p in sorted(it):
        if not p.is_file() or p.suffix.lower() not in exts:
            continue
        parts = {x.lower() for x in p.relative_to(d).parts[:-1]}
        if parts & {"output", "filechecks", ".git"}:
            continue
        out.append(p)
    return out


def fcf_profile(p: Path) -> dict:
    """讀一遍, 記下列數、欄名、鍵集合、期間分布。

    只對「票號+票聯」那兩欄算雜湊, 不是全部欄位。全欄位逐列接字串再雜湊,
    四十幾欄乘幾十萬列, 一個檔要二十幾秒, 而配對根本用不到那個精度。
    """
    cols = fc_cols({"files": [p]})
    keyi = fc_keycols(cols)
    pick = fc_pick_period(cols)
    keys: set = set()
    per: Counter = Counter()
    n = 0
    bad: list = []
    for tbl in fc_batches(p, bad):
        n += tbl.num_rows
        if keyi:
            kc = [cols[i] for i in keyi if cols[i] in tbl.column_names]
            if kc:
                keys.update(fc_hash(tbl, kc))
        if pick and pick in tbl.column_names:
            # 日期欄一個檔只有幾十到幾百個不同的值, 但有幾十萬列。先數出現
            # 次數再對「不同的值」跑 regex, 不要每一列都跑一次。
            import pyarrow.compute as pc
            for s in pc.value_counts(tbl.column(pick)):
                ym = _ym(s["values"].as_py())
                if ym:
                    per[ym] += s["counts"].as_py()
    return {"path": p, "name": p.name, "bytes": p.stat().st_size,
            "cols": cols, "rows": n, "keys": keys, "per": per,
            "haskey": bool(keyi), "bad": bad}


def fcf_stock(name: str) -> str:
    """檔名裡的票券系列標記。兩邊寫法不同但這個標記是一致的。"""
    k = name.lower()
    for tag in ("043", "160"):
        if tag in k:
            return tag
    return ""


def fcf_span(per: Counter) -> str:
    if not per:
        return "(看不出期間)"
    s = sorted(per)
    a, b = s[0], s[-1]
    t = f"{a[0]}-{a[1]:02d}"
    if a != b:
        t += f"..{b[0]}-{b[1]:02d}"
    return f"{t}  {len(per)} 個月"


FCF_CAP = 120_000        # 逐欄診斷最多留這麼多列, 再多記憶體吃不消


def fcf_pull(p: Path, cols: list, want: set, keyi: list, cap: int) -> dict:
    """再讀一次, 只留雜湊落在 want 裡的列, 用鍵索引起來。"""
    out: dict = {}
    n = 0
    for tbl in fc_batches(p, []):
        vals = None
        for r, h in enumerate(fc_hash(tbl, cols)):
            if h not in want:
                continue
            if vals is None:
                vals = [tbl.column(c).to_pylist() for c in cols]
            row = [("" if vals[k][r] is None else vals[k][r])
                   for k in range(len(cols))]
            out.setdefault(tuple(row[i] for i in keyi), []).append(row)
            n += 1
            if n >= cap:
                return out
    return out


def fcf_why(u: dict, c: dict, common: list, ca: Counter, cb: Counter) -> None:
    """兩邊都有的鍵, 到底是哪幾欄不一樣。

    只報「有 N 列不同」沒辦法判斷 —— 一欄的日期格式不同就會讓每一列都算
    不同, 跟真的少了資料是兩回事。
    """
    keyi = fc_keycols(common)
    if not keyi:
        add("         (找不到票號欄, 沒辦法逐欄比)")
        return
    only_a, only_b = set((ca - cb)), set((cb - ca))
    if len(only_a) > FCF_CAP:
        add(f"         差異超過 {FCF_CAP:,} 列, 只抽前面這些診斷")
    ra = fcf_pull(u["path"], common, only_a, keyi, FCF_CAP)
    rb = fcf_pull(c["path"], common, only_b, keyi, FCF_CAP)
    percol: Counter = Counter()
    order: Counter = Counter()          # 只是多值欄位的元素順序不同
    sample: dict = {}
    osample: dict = {}
    paired = lonely = same_but_order = 0
    for k, rows_a in ra.items():
        cand = rb.get(k)
        if not cand:
            lonely += 1
            continue
        paired += 1
        x, y = rows_a[0], cand[0]
        hit = ordhit = 0
        for i in range(len(common)):
            if x[i] == y[i]:
                continue
            # "a|b" 對 "b|a" 裝的是同一組值。GROUP_CONCAT 沒有 ORDER BY
            # 時元素順序是未定義的, 逐字比會全部比不中, 但那不是內容不同。
            if fc_same_set(x[i], y[i]):
                order[common[i]] += 1
                osample.setdefault(common[i], (x[i], y[i]))
                ordhit += 1
            else:
                percol[common[i]] += 1
                sample.setdefault(common[i], (x[i], y[i]))
                hit += 1
        if ordhit and not hit:
            same_but_order += 1
    if lonely:
        add(f"         其中 {lonely:,} 個票號在對方檔案裡根本沒有 "
            "-> 真的少了資料, 不是值不同")
        # 只說幾個沒用, 要追就得知道是哪幾張票
        miss = [k for k in ra if k not in rb]
        for k in miss[:10]:
            add(f"           票號 {' / '.join(str(x) for x in k)}")
        if len(miss) > 10:
            add(f"           ... 另外 {len(miss) - 10:,} 個")
    if not paired:
        return
    if same_but_order:
        add(f"         其中 {same_but_order:,} 列**內容相同, 只是多值欄位的"
            "元素順序不同**:")
        for col, n in order.most_common(3):
            a, b = osample[col]
            add(f"           {fc_cname(col):<22} {n:>7,} 列")
            add(f"             ita   ={str(a)[:64]}")
            add(f"             client={str(b)[:64]}")
        add("           -> 這不是差異。同一組值排列不同, 逐字比才會比不中。")
    if not percol:
        if not same_but_order:
            add("         (值都一樣, 差的是共同欄位以外的東西)")
        return
    add(f"         {sum(percol.values()):,} 處真的值不同, 差在這幾欄:")
    for col, n in percol.most_common(6):
        a, b = sample[col]
        add(f"           {fc_cname(col):<26} {n:>7,} 列"
            f"   ita={str(a)[:26]!r}  client={str(b)[:26]!r}")
    if len(percol) > 6:
        add(f"           ... 另外 {len(percol) - 6} 欄")


# --------------------------------------------------------------- --cpn
# after-ARA 的收取理論上是 coupon level: 一張票的一個航段收一次稅。同一個
# (票號, 票聯) 出現兩次, 要嘛是整列重複 (匯出重跑), 要嘛是同鍵不同值 (兩
# 筆對同一航段的不同記錄)。這兩種對計算的影響不一樣, 所以要分開數:
#
#   exact      —— 所有欄位一模一樣。去掉一份不影響任何金額。
#   business   —— 鍵相同但至少一欄不同。去掉哪一份會改變結果, 不能自動決定。
#
# 只報總數沒用, 要給得出「哪個檔、哪一列、差在哪一欄」才查得下去。
def cpn_key(cols: list) -> list:
    """(票號, 票聯) 的欄位位置。兩個都要有才算得上 coupon level。"""
    idx = []
    for want in (("tkt", "ticket"), ("coupon",)):
        for i, c in enumerate(cols):
            k = c.lower()
            if any(w in k for w in want) and ("num" in k or "no" in k
                                              or "coupon" in k):
                idx.append(i)
                break
    return idx if len(idx) == 2 else []


def mode_cpn(root: Path, where: str, show: int, quick: bool,
             nfile: int) -> None:
    d = root / where
    if not d.is_dir():
        sys.exit(f"找不到 {d}")
    files = fcf_files(d, (".csv", ".parquet"), deep=True)
    if not files:
        sys.exit(f"{d} 底下沒有 csv/parquet")
    alln = len(files)
    files = files[:nfile] if nfile > 0 else files

    add("")
    add("=" * 72)
    add(f"coupon level 重複檢查   {where}/   掃 {len(files)} / {alln} 個檔")
    add("=" * 72)
    add("  同一個 (票號, 票聯) 出現多次就是重複。分兩種:")
    add("    exact    整列一模一樣 —— 去掉一份不影響金額")
    add("    business 鍵相同但有欄位不同 —— 去掉哪一份會改變結果")
    add("")

    tot_rows = tot_keys = 0
    grand_e = grand_b = 0
    xl_sum: list = []
    xl_det: list = []
    for p in tqdm(files, desc="  掃描", unit="檔", leave=False):
        try:
            cols = fc_cols({"files": [p]})
        except Exception as e:
            add(f"  !! {p.name}  讀不到欄名: {type(e).__name__}")
            continue
        ki = cpn_key(cols)
        if not ki:
            add(f"  -- {p.name[:56]}  沒有票號+票聯欄, 跳過")
            continue
        # 第一遍: 每個鍵出現幾次, 以及整列雜湊有幾種
        seen: dict = {}
        n = 0
        bad: list = []
        cut = False
        try:
            for tbl in fc_batches(p, bad):
                n += tbl.num_rows
                kc = [cols[i] for i in ki]
                for k, h in zip(fc_joined(tbl, kc), fc_hash(tbl, cols)):
                    cur = seen.get(k)
                    if cur is None:
                        seen[k] = [1, h, True]   # 次數, 第一個雜湊, 是否全同
                    else:
                        cur[0] += 1
                        cur[2] = cur[2] and h == cur[1]
                if quick:
                    # 只要例子的話, 湊夠就停。這樣得到的是「掃到這裡為止」
                    # 的數字, 不是全檔統計 —— 輸出會講明, 不能當成總數。
                    d = [v for v in seen.values() if v[0] > 1]
                    if (sum(1 for v in d if v[2]) >= show
                            and sum(1 for v in d if not v[2]) >= show):
                        cut = True
                        break
        except Exception as e:
            add(f"  !! {p.name[:56]}  {type(e).__name__}: {str(e)[:60]}")
            continue
        dup = {k: v for k, v in seen.items() if v[0] > 1}
        tot_rows += n
        tot_keys += len(seen)
        if not dup:
            add(f"  OK {p.name[:56]}   {n:,} 列, 沒有重複的 coupon")
            continue
        exact = {k: v for k, v in dup.items() if v[2]}
        biz = {k: v for k, v in dup.items() if k not in exact}
        grand_e += sum(v[0] - 1 for v in exact.values())
        grand_b += sum(v[0] - 1 for v in biz.values())
        add("")
        add(f"  !! {p.name[:56]}")
        tail = "  [只掃到這裡就停了, 不是全檔統計]" if cut else ""
        add(f"       {n:,} 列 -> {len(seen):,} 個不同的 (票號, 票聯){tail}")
        add(f"       重複的鍵 {len(dup):,} 個 "
            f"(exact {len(exact):,} / business {len(biz):,})")
        add(f"       多出來的列: exact {sum(v[0]-1 for v in exact.values()):,}, "
            f"business {sum(v[0]-1 for v in biz.values()):,}")
        if bad:
            add(f"       (另有 {len(bad):,} 列讀不動, 沒算進去)")
        sm, dt = cpn_show(p, cols, ki, exact, biz, show)
        xl_sum += sm
        xl_det += dt

    add("")
    add("=" * 72)
    if quick:
        add("  (--quick: 每個檔湊夠例子就停, 下面的數字是抽樣不是總數)")
    add(f"合計   {tot_rows:,} 列 / {tot_keys:,} 個不同的 (票號, 票聯)")
    add(f"  exact 重複多出來的列    {grand_e:,}")
    add(f"  business 重複多出來的列 {grand_b:,}")
    if grand_b:
        add("")
        add("  business 那一批要人看過才能決定留哪一份 —— 兩列的金額或")
        add("  航段資訊不同, 自動去重等於替客戶做了判斷。")
    cpn_xlsx(root, xl_sum, xl_det)


def mode_cpnrows(root: Path, where: str) -> None:
    """把 cpn_examples.xlsx 裡那些票的原始列全部撈出來, 給客戶對。

    摘要那頁只說「差在 collect_tax」, 客戶要的是「我這一筆到底長怎樣、在
    哪個檔的第幾列」。所以這裡輸出的是原始整列加上位置。
    """
    import pandas as pd

    src = root / "cpn_examples.xlsx"
    if not src.is_file():
        sys.exit(f"找不到 {src} —— 先跑 python csvtools.py --cpn")
    want = pd.read_excel(src, "摘要", dtype=str)
    need = ("檔案", "類型", "票號", "票聯")
    miss = [c for c in need if c not in want.columns]
    if miss:
        sys.exit(f"{src} 的摘要頁少了欄位: {miss}")

    add("")
    add("=" * 72)
    add(f"把 {src.name} 裡的票撈出原始列   {len(want):,} 組")
    add("=" * 72)

    d = root / where
    byname: dict = {}
    for p in fcf_files(d, (".csv", ".parquet"), deep=True):
        byname.setdefault(p.name, p)

    out_rows: list = []
    for fname, grp in want.groupby("檔案"):
        p = byname.get(fname)
        if p is None:
            add(f"  !! {fname}  在 {where}/ 找不到, 跳過")
            continue
        keys = {(str(r["票號"]), str(r["票聯"])): str(r["類型"])
                for _, r in grp.iterrows()}
        cols = fc_cols({"files": [p]})
        ki = cpn_key(cols)
        if not ki:
            add(f"  !! {fname}  沒有票號+票聯欄, 跳過")
            continue
        kc = [cols[i] for i in ki]
        got = 0
        seen_at = 0
        import pyarrow as pa
        for tbl in fc_batches(p, []):
            ks = [tuple(x.split("\x1f")) for x in fc_joined(tbl, kc)]
            mask = [k in keys for k in ks]
            if any(mask):
                sub = tbl.filter(pa.array(mask))
                vals = [sub.column(c).to_pylist() for c in cols]
                hit = [(i, k) for i, k in enumerate(ks) if k in keys]
                for r, (off, k) in enumerate(hit):
                    # 原始列原封不動, 只在前面加檔名與類型兩欄
                    rec = {"檔案": fname, "類型": keys[k]}
                    for i, c in enumerate(cols):
                        rec[c if c.strip() else f"(空白欄{i})"] = (
                            "" if vals[i][r] is None else vals[i][r])
                    out_rows.append(rec)
                    got += 1
            seen_at += tbl.num_rows
        add(f"  {fname[:52]:<52} {len(keys):>3} 組 -> {got:>4} 列")

    if not out_rows:
        sys.exit("一列都沒撈到")
    # 不同檔的欄位集不一樣時, 併起來會有空洞。給客戶看的檔不要出現 NaN。
    df = pd.DataFrame(out_rows).fillna("")
    # 同一張票的幾筆要排在一起才好上下對照。用原始欄位排, 不另外加欄。
    by = ["檔案", "類型"] + [c for c in df.columns
                             if c.strip().lower() in ("tkt_num",
                                                      "coupon_number")]
    df = df.sort_values(by, kind="stable")
    out = root / "cpn_check.xlsx"
    if not cpn_write(out, [("明細", df)]):
        return
    add("")
    add(f"  寫到 {out.name}   {len(df):,} 列 x {len(df.columns)} 欄")
    add("  原始列原封不動, 只在最前面加了「檔案」與「類型」兩欄。")
    add("  同一張票的幾筆排在一起, 可以直接上下對照。")
    add("")
    add("  ! 這個檔有票號, 屬個人資料, 只走內部管道。")


def cpn_write(out: Path, sheets: list) -> bool:
    """寫 Excel, 並把票號那幾欄標成文字。回傳有沒有寫成功。"""
    import pandas as pd
    try:
        with pd.ExcelWriter(out, engine="openpyxl") as w:
            for name, df in sheets:
                df.to_excel(w, sheet_name=name, index=False)
    except Exception as e:
        add(f"  !! 寫不出 Excel ({type(e).__name__}: {e}), 改寫 CSV")
        for name, df in sheets:
            df.to_csv(out.with_name(f"{out.stem}_{name}.csv"), index=False,
                      encoding="utf-8-sig")
        return False
    # 票號的前導零是資料的一部分。存的已經是字串, 再把格式標成文字, 免得
    # 有人在 Excel 裡重存時被轉成數字。這一步失敗不能影響已經寫好的檔。
    try:
        import openpyxl
        wb = openpyxl.load_workbook(out)
        asnum = {"票號", "票聯", "tkt_num", "coupon_number"}
        for name, df in sheets:
            ws = wb[name]
            for ci, cname in enumerate(df.columns, 1):
                if str(cname).strip().lower() in asnum:
                    for r in range(2, ws.max_row + 1):
                        ws.cell(r, ci).number_format = "@"
        wb.save(out)
    except Exception as e:
        add(f"  (票號欄沒能標成文字格式: {type(e).__name__}; 值本身是對的)")
    return True


def cpn_xlsx(root: Path, summary: list, detail: list) -> None:
    import pandas as pd
    if not summary:
        add("")
        add("  沒有例子可以輸出。")
        return
    out = root / "cpn_examples.xlsx"
    df_s = pd.DataFrame(summary)
    df_d = pd.DataFrame(detail)
    if not cpn_write(out, [("摘要", df_s), ("明細", df_d)]):
        return
    add("")
    add(f"  例子寫到 {out.name}")
    add(f"    摘要  {len(df_s):,} 組重複 —— 每組一列, 有票號、票聯、出現次數、"
        "差在哪幾欄")
    add(f"    明細  {len(df_d):,} 列 —— 那幾組的原始整列, 五十幾欄都在, "
        "可以直接看")
    add("")
    add("  ! 這個檔有票號, 屬個人資料, 只走內部管道。")


def cpn_show(p: Path, cols: list, ki: list, exact: dict, biz: dict,
             show: int) -> None:
    """把例子的整列抓回來, 指出 business 那種差在哪一欄。"""
    want = set(list(exact)[:show]) | set(list(biz)[:show])
    if not want:
        return
    import pyarrow as pa
    got: dict = {}
    kc = [cols[i] for i in ki]
    for tbl in fc_batches(p, []):
        # 先把要的那幾列篩出來再轉 Python —— 原本是整批十萬列乘五十幾欄
        # 全部 to_pylist, 光為了抓三個例子, 一個檔要跑二三十秒。
        ks = fc_joined(tbl, kc)
        mask = [k in want for k in ks]
        if not any(mask):
            continue
        sub = tbl.filter(pa.array(mask))
        vals = [sub.column(c).to_pylist() for c in cols]
        for r, k in enumerate([k for k in ks if k in want]):
            got.setdefault(k, []).append(
                [("" if vals[i][r] is None else vals[i][r])
                 for i in range(len(cols))])
        if all(len(got.get(k, ())) > 1 for k in want):
            break                      # 每個要的鍵都湊齊兩列以上了, 不用再讀
    summary, detail = [], []
    for tag, grp in (("exact", exact), ("business", biz)):
        keys = [k for k in list(grp)[:show] if k in got]
        if not keys:
            continue
        add(f"       {tag} 的例子:")
        for k in keys:
            rows_ = got[k]
            tk = k.split("\x1f")
            add(f"         票號 {tk[0]}  票聯 {tk[-1]}   出現 {len(rows_)} 次")
            diff = [i for i in range(len(cols))
                    if len({r[i] for r in rows_}) > 1]
            for i in diff[:6]:
                vs = " | ".join(sorted({str(r[i])[:20] for r in rows_}))
                add(f"           {fc_cname(cols[i]):<24} {vs}")
            if len(diff) > 6:
                add(f"           ... 另外 {len(diff) - 6} 欄不同")
            summary.append({
                "檔案": p.name, "類型": tag, "票號": tk[0], "票聯": tk[-1],
                "出現次數": len(rows_),
                "不同的欄位數": len(diff),
                "不同的欄位": ", ".join(
                    (cols[i] or f"(空白欄{i})") for i in diff) or "(無)",
                "各欄的值": " ;; ".join(
                    f"{cols[i] or f'(空白欄{i})'}=" + "|".join(
                        sorted({str(r[i]) for r in rows_}))
                    for i in diff[:6]) or "(整列一模一樣)",
            })
            for seq, r in enumerate(rows_, 1):
                rec = {"檔案": p.name, "類型": tag, "票號": tk[0],
                       "票聯": tk[-1], "第幾筆": seq}
                for i, c in enumerate(cols):
                    rec[c if c.strip() else f"(空白欄{i})"] = r[i]
                detail.append(rec)
    return summary, detail


def mode_fc_flat(root: Path, rows: bool) -> None:
    base = root / FC
    if not base.is_dir():
        sys.exit(f"找不到 {base}")

    def sub(*names):
        for d in base.iterdir():
            if d.is_dir() and d.name.lower() in names:
                return d
        return None

    di, dc = sub("ita"), sub("client", "kpmg", "ours")
    ita_f = (fcf_files(di, (".csv", ".txt", ".parquet"), deep=True) if di
             else fcf_files(base, (".csv", ".txt", ".parquet"), deep=False))
    if not ita_f:
        sys.exit(f"{base} 底下找不到 ita 的檔")

    # client 側: filechecks/client/ 全收 (那是人挑好的), 再從 data/ 補進
    # 期間可能重疊的。data/ 有三百多個檔三十幾 GB, 全讀要兩個多小時, 但大
    # 部分月份跟這批 ita 根本不相干。
    #
    # 檔名只決定「值不值得讀」, 不決定「是不是同一份」—— 那還是看內容。
    # 檔名判斷錯最多是多讀或漏讀一個候選, 不會讓配對結果錯。
    want: set = set()
    for p in ita_f:
        sp, _ = parse_span(p.name)
        if sp:
            want |= sp
    # client/ 是人挑好的, 但裡面可能混著別的階段 (ARA 那批是 2022-2025,
    # 跟這次的 2014-2021 一個月都不重疊, 讀了也只是白讀)。一樣用期間濾,
    # 但濾掉哪些要印出來 —— 人放進去的東西不能無聲跳過。
    picked, dropped = [], []
    for p in (fcf_files(dc, (".csv", ".txt", ".parquet"), deep=True) if dc else []):
        sp, _ = parse_span(p.name)
        if sp is None or sp & want:
            picked.append(p)
        else:
            dropped.append(p)
    extra, skipped, blind = [], 0, []
    dat = root / DATA_DIR
    if dat.is_dir():
        have = {q.name for q in picked}
        for p in fcf_files(dat, (".csv", ".parquet"), deep=True):
            if p.name in have:
                continue
            sp, _ = parse_span(p.name)
            if sp is None:
                blind.append(p)
            elif sp & want:
                extra.append(p)
            else:
                skipped += 1
    cli_f = picked + extra + blind
    if not cli_f:
        sys.exit("找不到 client 側的檔")
    src = (f"ita {di.name if di else base.name}/   client "
           + (f"{dc.name}/ {len(picked)} 檔" if dc else "")
           + (f" + {DATA_DIR}/ 期間相符 {len(extra)} 檔" if extra else "")
           + (f" + 期間看不出來 {len(blind)} 檔" if blind else "")
           + (f"   ({skipped} 個因期間不重疊沒讀)" if skipped else ""))

    add("")
    add("=" * 72)
    add(f"ita vs client 內容配對   ita {len(ita_f)} 檔 / client {len(cli_f)} 檔")
    add("=" * 72)
    add(f"  {src}")
    add("  配對只看內容 (票號+票聯的重疊), 不看檔名 —— 檔名兩邊的寫法不一樣。")
    if dropped:
        add("")
        add(f"  client/ 裡有 {len(dropped)} 個檔跟這批 ita 一個月都不重疊, 沒讀:")
        for p in dropped[:12]:
            sp, _ = parse_span(p.name)
            l = sorted(sp or [])
            r = (f"{l[0][0]}-{l[0][1]:02d}..{l[-1][0]}-{l[-1][1]:02d}"
                 if len(l) > 1 else (f"{l[0][0]}-{l[0][1]:02d}" if l else "?"))
            add(f"      {p.name[:52]:<52} {r}")
        if len(dropped) > 12:
            add(f"      ... 另外 {len(dropped) - 12} 個")

    # 先用檔案大小分組, 只有大小一樣的才值得算 SHA256。內容相同的 CSV 是
    # 位元組相同, 這比逐列算指紋快幾個數量級。
    bysize: dict = {}
    for p in ita_f:
        bysize.setdefault(p.stat().st_size, []).append(p)
    byte_dup = []
    for sz, ps in bysize.items():
        if len(ps) < 2:
            continue
        h: dict = {}
        for p in ps:
            h.setdefault(sha256(p), []).append(p)
        byte_dup += [v for v in h.values() if len(v) > 1]

    bar = tqdm(total=len(ita_f) + len(cli_f), desc="  讀取", unit="檔",
               leave=False)
    # 有一個檔讀不動不該讓整輪白跑 —— 這批裡有 .txt, 分隔符不一定是逗號
    ita, cli, dead = [], [], []
    for side, src in ((ita, ita_f), (cli, cli_f)):
        for p in src:
            try:
                side.append(fcf_profile(p))
            except Exception as e:
                dead.append((p, f"{type(e).__name__}: {e}"))
            bar.update(1)
    bar.close()
    if dead:
        add("")
        add("  讀不動的檔 (跳過, 其餘照比):")
        for p, why in dead:
            add(f"    !! {p.name}   {why[:90]}")
    if not ita:
        sys.exit("ita 側一個檔都讀不動")
    broken = [u for u in ita + cli if u.get("bad")]
    if broken:
        add("")
        add("  有壞列的檔 (壞的那幾列跳過, 其餘照讀):")
        for u in broken:
            b = u["bad"]
            add(f"    !! {u['name'][:56]}   跳過 {len(b):,} 列 "
                f"/ 讀進 {u['rows']:,} 列")
            for num, nc, txt in b[:3]:
                where = f"第 {num} 列" if num else "位置不明"
                add(f"         {where}: {nc} 欄  {txt}")
            if len(b) > 3:
                add(f"         ... 另外 {len(b) - 3:,} 列")
        add("      -> 跳過的列不參與比對, 底稿要寫明。")

    # 1. ita 自己有沒有重複的檔
    add("")
    add("-" * 72)
    add("ita 這批裡面有沒有同一份檔出現兩次")
    idx = {u["path"]: i for i, u in enumerate(ita)}
    dups = [[idx[p] for p in g if p in idx] for g in byte_dup]
    dups = [g for g in dups if len(g) > 1]
    kind = {tuple(sorted(g)): "位元組完全一樣" for g in dups}
    # 位元組不同但票號完全同一組的也要抓 —— 同一批列重新排序過就是這樣:
    # 檔案大小一模一樣, SHA256 卻不同, 只比位元組會漏掉。
    byrow: dict = {}
    for i, u in enumerate(ita):
        if u["keys"]:
            byrow.setdefault((u["rows"], len(u["keys"])), []).append(i)
    for g in byrow.values():
        for a in range(len(g)):
            for b in range(a + 1, len(g)):
                k = (g[a], g[b])
                if k in kind or ita[g[a]]["keys"] != ita[g[b]]["keys"]:
                    continue
                kind[k] = "票號完全同一組 (位元組不同, 可能只是列的順序不同)"
                dups.append(list(k))
    if dups:
        for g in dups:
            a, b = ita[g[0]], ita[g[1]]
            add(f"  !! 內容完全一樣:")
            add(f"       {a['name']}   {a['bytes']:,} bytes")
            add(f"       {b['name']}   {b['bytes']:,} bytes")
            add(f"       {a['rows']:,} 列, "
                + kind.get(tuple(sorted(g)), "內容相同"))
            add(f"       期間 {fcf_span(a['per'])}")
            if len(g) > 2:
                add(f"       (這一組共 {len(g)} 份)")
            add("     -> 其中一份標錯了, 要回頭問 ita 哪一份才是對的。")
    else:
        add("  沒有重複。")

    # 2. 配對
    add("")
    add("-" * 72)
    add("配對")
    # 期間先當硬性條件: 兩邊都看得出月份, 卻一個月都不重疊, 就不可能是同一
    # 份, 不管鍵重疊多少。票號是流水號, 不同月份撞在一起很正常, 光看重疊率
    # 會配出 2017 年對 2014 年這種東西。
    cand = []
    for i, u in enumerate(ita):
        for j, c in enumerate(cli):
            if not u["keys"] or not c["keys"]:
                continue
            if u["per"] and c["per"] and not (set(u["per"]) & set(c["per"])):
                continue
            ov = len(u["keys"] & c["keys"])
            # 幾十萬個票號裡中兩個是雜訊, 不是配對。低於 1% 就當沒對手 ——
            # 報一個 0.00% 的「對手」比報「沒有對手」更容易讓人看漏。
            if ov < 50 and ov < 0.01 * len(u["keys"]):
                continue
            # 先看「ita 有多少被涵蓋」, 同分看是不是同一個系列, 再同分才看
            # Jaccard。要問的是「ita 這份的內容在不在我們手上」, 不是兩份
            # 大小像不像。
            #
            # 系列這一關是必要的: 160 stock 的月檔跟「整個月的 KA」票號會
            # 全部對得上 (前者是後者的子集), 光看重疊會配到後者去, 然後每
            # 一列都比出差異。系列只在同分時當決勝, 不會凌駕內容。
            cand.append((round(ov / len(u["keys"]), 6),
                         1 if fcf_stock(u["name"]) == fcf_stock(c["name"]) else 0,
                         ov / len(u["keys"] | c["keys"]), ov, i, j))
    # 全域由高到低指派, 不是每個 ita 各自搶 —— 各自搶會先來先贏, 後面連環錯
    cand.sort(reverse=True)
    take_i: dict = {}
    take_j: dict = {}
    for cov, _same, jac, ov, i, j in cand:
        if i in take_i:
            continue
        # 一個 client 檔可以被好幾個 ita 檔認領 —— ita 送的是月檔, 我們手上
        # 是年檔, 十二個月本來就該對到同一個檔。但認領它的那幾份彼此不能有
        # 重疊的票號: 有重疊就不是「不同月份」, 是同一份資料出現兩次。
        if any(ita[i]["keys"] & ita[k]["keys"] for k in take_j.get(j, ())):
            continue
        take_i[i] = (j, cov, ov)
        take_j.setdefault(j, []).append(i)
    # 內容一模一樣的 ita 檔只會有一份搶到對手 (誰搶到是任意的)。同一組的
    # 其他份指向同一個 client, 並記下是跟誰共用, 不然會被誤報成「沒有對手」。
    # 內容相同的幾份, 答案一定要一樣。誰先搶到是任意的, 而且票號重疊會讓
    # 後面幾份被擋掉、掉到某個爛對手上 (實測出現過 0.00% 的配對)。所以整組
    # 統一跟涵蓋率最高的那一份走。
    shared: dict = {}
    for g in dups:
        got = [i for i in g if i in take_i]
        if not got:
            continue
        best = max(got, key=lambda i: take_i[i][1])
        j, cov, ov = take_i[best]
        for i in g:
            if i != best:
                take_i[i] = (j, cov, ov)
                shared[i] = best

    # 沒配到的要講得出為什麼: 是完全沒有重疊, 還是有重疊但期間對不上
    blocked: dict = {}
    for i, u in enumerate(ita):
        if i in take_i or not u["keys"]:
            continue
        best = (0, None, None)
        for j, c in enumerate(cli):
            if not c["keys"]:
                continue
            ov = len(u["keys"] & c["keys"])
            if ov > best[0]:
                # 那個對手是不是已經被另一份 ita 檔認領了? 是的話要講出來,
                # 不能含糊帶過成「期間對不上」—— 期間其實一樣。
                rival = next((k for k in take_j.get(j, ())
                              if ita[k]["keys"] & u["keys"]), None)
                best = (ov, c, rival)
        if best[0]:
            blocked[i] = best

    used = set(take_j)
    plan = []
    for i, u in enumerate(ita):
        if i in take_i:
            j, cov, _ = take_i[i]
            plan.append((u, cli[j], cov, shared.get(i)))
        else:
            plan.append((u, None, blocked.get(i), None))

    for u, c, s, sh in sorted(plan, key=lambda x: x[0]["name"]):
        add("")
        add(f"  {u['name']}")
        add(f"      ita     {u['rows']:>10,} 列  {u['bytes']:>13,} bytes  "
            f"{fcf_span(u['per'])}")
        if c is None:
            if not u["haskey"]:
                add("      !! 這個檔沒有票號欄, 配不了")
            elif s:
                ov, cc, rival = s
                pct = ov / max(len(u["keys"]), 1)
                add(f"      !! 沒有對手。最接近的是 {cc['path'].name}")
                add(f"         票號重疊 {ov:,} / ita {len(u['keys']):,} "
                    f"= {pct:.2%}")
                if rival is not None:
                    add(f"         那一份已經被 {ita[rival]['name']} 認領, "
                        "而你們兩份的票號互相重疊 ——")
                    add("         代表這兩個 ita 檔裝的是同一批資料, "
                        "不是兩段不同期間。")
                else:
                    add(f"         但期間對不上 (對方 {fcf_span(cc['per'])}) "
                        "—— 票號是流水號, 光重疊不算數")
            else:
                add("      !! 沒有任何 client 檔跟它有票號重疊")
            continue
        ov = len(u["keys"] & c["keys"])
        add(f"      client  {c['rows']:>10,} 列  {c['bytes']:>13,} bytes  "
            f"{fcf_span(c['per'])}")
        add(f"              {c['path'].relative_to(root)}")
        mark = "" if s > 0.99 else "   <-- 不是全部對得上"
        add(f"      重疊 {ov:,} / ita {len(u['keys']):,} 個鍵 = {s:.2%}{mark}")
        if sh is not None:
            add(f"      註: 跟 {ita[sh]['name']} 內容完全相同, 共用同一個對手")
        if u["rows"] != c["rows"]:
            add(f"      !! 列數差 {abs(u['rows'] - c['rows']):,}")
        only = [x for x in u["cols"] if x not in c["cols"]]
        if only:
            add(f"      ita 多出的欄: {', '.join(only[:6])}")

    left = [c for i, c in enumerate(cli) if i not in used]
    add("")
    add(f"  client 有 {len(left)} 個檔這次沒有對手 (ita 沒送這些月份, 正常)")
    if not rows:
        add("")
        add("  配對看起來對的話, 再跑內容比對:")
        add("      python csvtools.py --fc --flat --rows")
        return

    # ---- 內容比對 ----
    add("")
    add("-" * 72)
    add("內容比對   共同欄位上逐列比, 順序不算數")
    ok = bad = 0
    for u, c, s, sh in sorted(plan, key=lambda x: x[0]["name"]):
        if c is None:
            continue
        common = [x for x in u["cols"] if x in c["cols"]]
        if not common:
            add(f"  !! {u['name']}  沒有共同欄位, 比不了")
            bad += 1
            continue
        # 這裡一樣要帶 bad —— 不帶的話壞列會在做完一堆工之後才把整輪炸掉
        ba: list = []
        ca: Counter = Counter()
        for tbl in fc_batches(u["path"], ba):
            ca.update(fc_hash(tbl, common))
        bb: list = []
        cb: Counter = Counter()
        for tbl in fc_batches(c["path"], bb):
            cb.update(fc_hash(tbl, common))
        only_a = sum((ca - cb).values())
        only_b = sum((cb - ca).values())
        add("")
        add(f"  {u['name']}")
        add(f"      vs {c['path'].relative_to(root)}")
        add(f"      共同欄位 {len(common)} / ita {len(u['cols'])} 欄, "
            f"client {len(c['cols'])} 欄")
        miss = [x for x in u["cols"] if x not in c["cols"]]
        extra = [x for x in c["cols"] if x not in u["cols"]]
        if miss:
            add(f"      只有 ita 有的欄: {', '.join(miss[:8])}")
        if extra:
            add(f"      只有 client 有的欄: {', '.join(extra[:8])}")
        if ba or bb:
            add(f"      !! 跳過壞列: ita {len(ba):,} 列, client {len(bb):,} 列 "
                "(沒參與比對)")
        if not only_a:
            # ita 每一列在 client 都找得到。多出來的是年度/季檔裡其他月份,
            # 不是差異 —— 這一題問的是「ita 送來的東西我們有沒有」。
            ok += 1
            if only_b:
                add(f"      一致 —— ita {u['rows']:,} 列全部在對方檔案裡")
                add(f"         (對方多 {only_b:,} 列, 是同一個檔涵蓋的其他月份)")
            else:
                add(f"      一致 —— {u['rows']:,} 列在共同欄位上完全相同")
        else:
            bad += 1
            add(f"      !! 只在 ita 有 {only_a:,} 列   "
                f"只在 client 有 {only_b:,} 列")
            fcf_why(u, c, common, ca, cb)
    add("")
    add(f"  一致 {ok} 組   有差異 {bad} 組")
    add("  (「一致」含 ita 是對方子集的情形 —— ita 每一列對方都有就算過)")
    if bad == 0 and ok:
        add("  配對到的全部一致 —— ita 這批跟我們手上的內容相同。")


def mode_fc(root: Path, rows: bool, loose: bool = False,
            only: str | None = None, degrade: bool = False,
            mask: bool = False) -> None:
    try:
        import pyarrow  # noqa: F401
    except ImportError:
        sys.exit("要先裝 pyarrow:  pip install pyarrow")

    base = root / FC
    if not base.is_dir():
        sys.exit(f"找不到 {base}")

    # ita/<階段>/  對上  client/<階段>/
    stages = []
    for d in sorted((base / "ita").iterdir()) if (base / "ita").is_dir() else []:
        if d.is_dir() and (base / "client" / d.name).is_dir():
            stages.append(d.name)
    if not stages:
        sys.exit(f"{base} 底下要有 ita/<階段>/ 與 client/<階段>/ 兩組同名資料夾")

    add("")
    add("=" * 72)
    add("ita vs client 逐檔比對   " + ("內容比對" if rows else "配對計畫 (加 --rows 才真的比內容)"))
    add("=" * 72)

    plan = []
    for st in stages:
        ita = fc_units(base / "ita" / st)
        cli = fc_units(base / "client" / st)
        pairs, no_cli, no_ita, vote = fc_pair(ita, cli)
        plan.append((st, pairs, no_cli, no_ita, ita, cli, vote))

        add("")
        add(f"  [{st}]")
        add(f"    ita     {sum(len(u['files']) for u in ita):>3} 檔 -> "
            f"{len(ita):>2} 個單位    " + " / ".join(sorted({u["series"] for u in ita})))
        add(f"    client  {sum(len(u['files']) for u in cli):>3} 檔 -> "
            f"{len(cli):>2} 個單位    " + " / ".join(sorted({u["series"] for u in cli})))
        add(f"    配對 {len(pairs)} 組"
            + (f"   ita 有 {len(no_cli)} 份找不到對手" if no_cli else "")
            + (f"   client 有 {len(no_ita)} 份沒人認領" if no_ita else ""))

        w = max([len(fc_label(a)) for a, _, _ in pairs] or [8])
        # 要拿 ita 那份去替換 client 那份來重跑時, 期間一致才換得動。對方
        # 那份涵蓋更多月份的話, 整檔換掉等於把其他月份一起刪了。
        ok_sw = [x for x in pairs if x[2] == "同期"]
        no_sw = [x for x in pairs if x[2] != "同期"]
        if not rows:   # --rows 底下每一對都會再列一次, 這裡就不重複了
            for a, b, how in sorted(pairs, key=_fc_sort):
                if how == "同期":
                    mark = "   [可直接換]"
                else:
                    gap = len(b["span"]) - len(a["span"])
                    mark = f"   !! 不能直接換 —— 對方那份多涵蓋 {gap} 個月"
                add(f"      {fc_label(a):<{w}}  {fc_names(a)}")
                add(f"      {'':<{w}}  {fc_names(b)}{mark}")
        add(f"    替換評估: {len(ok_sw)} 組期間一致可直接換, "
            f"{len(no_sw)} 組對方涵蓋更廣, 整檔換掉會少資料")
        for u in no_cli:
            add(f"      !! ita 這份在 client 找不到對手:  {fc_names(u)}  [{fc_label(u)}]")
        for u in no_ita:
            add(f"      !! client 這份沒有人認領:  {fc_names(u)}  [{fc_label(u)}]")
        # 系列對應是自己從期間推出來的, 推錯了底下整段比對都白做, 印出來讓
        # 人看一眼。兩邊名字不一樣不一定是錯 (命名規則本來就不同), 但值得停
        # 下來確認。
        used_map = {}
        for a, b, _ in pairs:
            used_map.setdefault(a["series"], set()).add(b["series"])
        if used_map:
            add("    系列對應")
            for k, v in sorted(used_map.items()):
                for t in sorted(v):
                    add(f"      [{k}]  ->  [{t}]"
                        + ("" if k == t else "   << 名字不一樣, 確認是同一批"))

    # ---- 欄位對照。欄位不一樣就沒必要比內容, 先看這個。
    add("")
    add("  欄位對照")
    seen = set()
    for st, pairs, _, _, _, _, _ in plan:
        for a, b, _ in pairs:
            k = (a["series"], b["series"])
            if k in seen:
                continue
            seen.add(k)
            ca, cb = fc_cols(a), fc_cols(b)
            sa, sb = set(ca), set(cb)
            add(f"    [{st}] {a['series']}  ->  {b['series']}")
            add(f"       ita {len(ca)} 欄   client {len(cb)} 欄   "
                + ("欄名完全一樣" if sa == sb else "!! 欄名不一樣"))
            if sa != sb:
                if sa - sb:
                    add(f"       只有 ita 有 ({len(sa - sb)}): "
                        + ", ".join(sorted(sa - sb)[:12]))
                if sb - sa:
                    add(f"       只有 client 有 ({len(sb - sa)}): "
                        + ", ".join(sorted(sb - sa)[:12]))
            elif ca != cb:
                add("       (欄位順序不同, 比對時會按欄名排序, 不影響)")

    # ---- 寫法對照。欄名一樣不代表值的寫法一樣: 日期可能一邊是 2017-08-01
    # 一邊是 01/08/2017, 票號可能一邊留著前導零一邊掉了。這種差異會讓每一
    # 列都比不中, 但原因不是資料不同。兩邊列序本來就不一樣, 所以不能比「第
    # 幾列的值」, 要比值的字元形狀 —— 數字換成 9、字母換成 A 之後的樣子。
    add("")
    add("  寫法對照 —— 每一組取開頭幾千列, 數字換 9 字母換 A 之後比形狀")
    nbad = 0
    shape_bad: set = set()
    for st, pairs, _, _, _, _, _ in plan:
        for a, b, _ in sorted(pairs, key=_fc_sort):
            cols = [c for c in fc_cols(a) if c in set(fc_cols(b))]
            sa = fc_shapes(a["files"][0], cols)
            sb = fc_shapes(b["files"][0], cols)
            diff = fc_shape_diff(sa, sb, cols)
            if not diff:
                continue
            nbad += 1
            shape_bad.add((st, fc_label(a)))
            add(f"    [{st}] {fc_label(a)}   {fc_names(a)}")
            w = min(26, max(len(c) for c, _, _, _ in diff))
            for c, xa, xb, why in diff[:8]:
                add(f"       {c[:w]:<{w}}  ita {xa[:24]:<24}  "
                    f"client {xb[:24]:<24}")
                add(f"       {'':<{w}}  !! {why}")
            if len(diff) > 8:
                add(f"       … 另外 {len(diff) - 8} 欄也一樣")
    if not nbad:
        add("    每一組的每一欄寫法都一致")
    else:
        add(f"    !! {nbad} 組的寫法對不上 —— 這種差異逐列比一定全軍覆沒, "
            "但原因是檔案格式, 不是資料")

    if not rows:
        add("")
        add("  上面沒問題就跑:  python csvtools.py --fc --rows")
        return

    # ---- 逐行比對
    tot = sum(a["bytes"] + b["bytes"]
              for _, pairs, _, _, _, _, _ in plan for a, b, _ in pairs)
    bar = tqdm(total=tot, unit="B", unit_scale=True, desc="  比對", leave=False)
    out, diffs, bad = [], [], 0
    loose_hit: set = set()
    deg_hit: dict = {}
    for st, pairs, _, _, _, _, _ in plan:
        ps = sorted(pairs, key=_fc_sort)
        if only:
            keys = [k.strip().lower() for k in only.split(",") if k.strip()]
            ps = [x for x in ps
                  if any(k in x[0]["files"][0].name.lower() for k in keys)]
        if not ps:
            continue
        # 檔名前面那一長串每個檔都一樣, 佔位置又看不出差別, 去掉。
        pre = os.path.commonprefix([a["files"][0].name for a, _, _ in ps])
        pre = pre[:pre.rfind(" ") + 1] if " " in pre else ""
        out.append((None, f"    [{st}]"))
        wl = max(len(fc_label(a)) for a, _, _ in ps)
        wn = max(len(fc_short(a["files"][0].name, pre)) for a, _, _ in ps)
        for a, b, how in ps:
            cols = sorted(set(fc_cols(a)) & set(fc_cols(b)))
            lo = (fc_multi(a, cols) | fc_multi(b, cols)) if loose else None
            dg = (fc_degrade_plan(a["files"][0], b["files"][0], cols)
                  if degrade else None)
            # 標了 both 的欄位兩側都要套, 其餘只套在 client 那一側
            dga = {k: v for k, v in dg.items() if v[2]} if dg else None
            if dg:
                deg_hit[fc_label(a)] = dg
            if lo:
                loose_hit |= lo
            # 兩邊各記各的, 不要相減成一個數。淨值只告訴你「我方多一列」,
            # 分不出是對方真的沒有這一列, 還是同一列我方存了兩份。
            ca: Counter = Counter()
            cb: Counter = Counter()
            na = nb = 0
            for f in a["files"]:
                for t in fc_batches(f):
                    na += t.num_rows
                    for h in fc_hash(t, cols, lo, dga):
                        ca[h] += 1
                bar.update(f.stat().st_size)
            for f in b["files"]:
                for t in fc_batches(f):
                    nb += t.num_rows
                    for h in fc_hash(t, cols, lo, dg):
                        cb[h] += 1
                bar.update(f.stat().st_size)
            ha = {k for k, v in ca.items() if v > cb.get(k, 0)}
            hb = {k for k, v in cb.items() if v > ca.get(k, 0)}
            only_a = sum(ca[k] - cb.get(k, 0) for k in ha)
            only_b = sum(cb[k] - ca.get(k, 0) for k in hb)
            dup_a = sum(ca[k] - cb[k] for k in ha if cb.get(k, 0))
            # 對方一份都沒有, 但我方存了好幾份 —— 那是我方檔內的重複列,
            # 跟「對方漏了一列資料」是兩回事, 結論寫法完全不同。
            self_dup = sum(ca[k] - 1 for k in ha if ca[k] > 1)
            if how == "同期":
                ok = not ha and not hb
                note = ("一致" if ok else
                        f"!! ita 多 {only_a:,}   client 多 {only_b:,}")
            else:
                # 對方那份涵蓋更多月份, 本來就會多。要證明的是「我方拿出來的
                # 每一列對方都有」—— 包含關係, 不是相等。
                ok = not ha
                note = (f"對方都有 (對方另有 {only_b:,} 列屬其他期間)"
                        if ok else f"!! ita 有 {only_a:,} 列對方沒有")
            bad += 0 if ok else 1
            out.append((ok, f"      {fc_label(a):<{wl}}  "
                            f"{fc_short(a['files'][0].name, pre):<{wn}}  "
                            f"ita {na:>9,}   client {nb:>9,}   {note}"))
            # ---- 對不上就要講出為什麼, 不然一個數字沒法追
            pad = " " * 10 + "-> "
            if not ok:
                # 對方涵蓋整年時, 「多出來」的列大半是別的月份, 要先篩掉,
                # 不然診斷會被無關的列淹掉。
                pc = fc_pick_period(cols)
                ymf = ((cols.index(pc), sorted(a["span"])[0])
                       if how != "同期" and pc in cols else None)
                if dup_a:
                    out.append((ok, pad + f"其中 {dup_a:,} 列對方其實也有, "
                                          "只是我方存了比較多份 (重複列)"))
                if self_dup:
                    out.append((ok, pad + f"其中 {self_dup:,} 列是 ita 檔內"
                                "自己的重複列 (同一列出現多次)"
                                + ("  —— 但這一組的寫法本來就對不上, 精度掉"
                                   "了之後不同的列會塌成一樣, 這個數字不算數"
                                   if (st, fc_label(a)) in shape_bad else "")))
                ra = fc_rowvals(a, cols, ha, FC_KEY_CAP, None, lo, dga)
                rb = fc_rowvals(b, cols, hb, FC_KEY_CAP, ymf, lo, dg)
                hint = None
                if ra and rb and len(ra) <= FC_DIAG_CAP and len(rb) <= FC_DIAG_CAP:
                    hits = fc_explain(cols, ra, rb)
                    if hits:
                        n_, i_, c_, va, vb = hits[0]
                        # cols 是按欄名排序過的, 這裡要換回檔案裡的位置
                        try:
                            pos = fc_cols(a).index(c_) + 1
                        except ValueError:
                            pos = i_ + 1
                        hint = (f"拿掉第 {pos} 欄 {fc_cname(c_)} 之後 "
                                f"{n_:,}/{min(len(ra), len(rb)):,} 列對上")
                        out.append((ok, pad + hint))
                        out.append((ok, pad + f"   ita {va[:40]!r}   "
                                              f"client {vb[:40]!r}"
                                    + ("   << 同一組值, 只是順序不同"
                                       if fc_same_set(va, vb) else "")))
                elif ra and rb:
                    ki = fc_keycols(cols)
                    if ki:
                        pr, nk, percol, sm = fc_bykey(cols, ra, rb, ki)
                        out.append((ok, pad + "用 "
                                    + "+".join(fc_cname(cols[i]) for i in ki)
                                    + f" 湊對: {pr:,} 列對方也有 (值不同), "
                                      f"{nk:,} 列對方連這個鍵都沒有"))
                        for c_, n_ in percol.most_common(3):
                            va, vb = sm[c_]
                            out.append((ok, pad + f"   {fc_cname(c_)} 有 "
                                                  f"{n_:,} 列不同   "
                                                  f"ita {va[:30]!r}  "
                                                  f"client {vb[:30]!r}"
                                        + ("   << 同一組值, 只是順序不同"
                                           if fc_same_set(va, vb) else "")))
                        if not pr:
                            # 鍵一個都對不上時, 看每一欄各自重疊多少 —— 全部
                            # 欄位都不重疊才是兩批資料, 只有幾欄不重疊代表
                            # 是那幾欄的寫法問題。
                            ov = [(r_, c_) for r_, c_ in
                                  fc_overlap(cols, ra, rb) if r_ < 0.999]
                            for (c_, s1, n1, x1, s2, n2, x2) in fc_colcompare(
                                    cols, ra, rb, [c for _, c in ov[:4]]):
                                out.append((ok, pad + f"{fc_cname(c_)}"))
                                out.append((ok, pad + f"   ita    {s1:<26} "
                                                      f"{n1[:24]} ~ {x1[:24]}"))
                                out.append((ok, pad + f"   client {s2:<26} "
                                                      f"{n2[:24]} ~ {x2[:24]}"))
                            out.append((ok, pad + (
                                f"逐欄重疊率: {len(cols) - len(ov)}/{len(cols)}"
                                " 欄的值兩邊完全重疊; 不重疊的是 "
                                + ", ".join(f"{fc_cname(c_)} {r_:.0%}"
                                            for r_, c_ in ov[:6])
                                + (" …" if len(ov) > 6 else "")
                                if ov else "逐欄重疊率: 每一欄的值兩邊都完全"
                                           "重疊, 是欄位的組合對不上")))
                    else:
                        out.append((ok, pad + f"對不上的列太多 (ita {len(ra):,},"
                                    f" client {len(rb):,}), 找不到可以湊對的"
                                    "鍵, 只能看 " + FC_DIFF))
                    hint = hint or (pr if ki else None)
                elif ra and not rb and dup_a < only_a:
                    # 對方那邊一列都沒剩 = 我方完整包含對方, 只是多出幾列。
                    # 不是「對方整批缺資料」—— 底下並排就會看到對方其實有
                    # 同一張票的別的列。
                    out.append((ok, pad + "對方那邊沒有任何一列是我方沒有的"
                                          " —— 我方是單方面多出來, 不是互有"
                                          "出入"))
                    hint = True
                # 只差幾列的時候, 與其叫人去翻 csv, 不如直接印出來。票號
                # 遮掉中間, 貼回來討論不會外流。
                if ra and only_a <= FC_SHOW:
                    out.append((ok, pad + f"這 {len(ra[:FC_SHOW])} 列長這樣"
                                + ("(值是降階後的)" if dg else "")
                                + (" (票號換成編號, 同號 = 同一張票)"
                                   if mask else "")
                                + ":"))
                    for r in ra[:FC_SHOW]:
                        fl = [f"{c}={fc_mask(c, v, not mask)}"
                              for c, v in zip(cols, r) if v]
                        bl = [c for c, v in zip(cols, r) if not v]
                        ln = ""
                        for piece in fl:
                            if len(ln) + len(piece) > 92:
                                out.append((ok, pad + "   " + ln))
                                ln = ""
                            ln += piece + "  "
                        if ln:
                            out.append((ok, pad + "   " + ln))
                        if bl:
                            out.append((ok, pad + f"      (這列有 {len(bl)} 欄"
                                        "是空的: " + ", ".join(bl[:6])
                                        + (" …" if len(bl) > 6 else "") + ")"))
                        out.append((ok, pad + "   " + "-" * 40))
                    # 把同一張票在兩邊的所有列並排 —— 只看「對方有沒有這
                    # 一列」不夠: 對方可能有同一張票的別的列。整組攤開才看
                    # 得出到底是多了一列、還是某一欄不同。
                    ki2 = fc_keycols(cols)
                    # 鍵本身被降階過就不能拿來分組 —— 票號降到 6 位有效數
                    # 字之後, 幾百張不同的票會共用同一個值, 並排出來的是一
                    # 堆不相干的列, 看了只會誤判。
                    lossy = [cols[i] for i in ki2
                             if dg and cols[i] in dg and dg[cols[i]][2]] if ki2 else []
                    if lossy:
                        out.append((ok, pad + "不並排同一張票了 —— "
                                    + ", ".join(fc_cname(c) for c in lossy)
                                    + " 本身已被降階, 拿來分組會把不同的票"
                                      "混在一起"))
                    elif ki2:
                        want_k = {tuple(r[i] for i in ki2) for r in ra[:FC_SHOW]}
                        ga = fc_findkey(a, cols, ki2, want_k, 400, None, lo, dga)
                        gb = fc_findkey(b, cols, ki2, want_k, 400, ymf, lo, dg)
                        miss = Counter(tuple(r) for r in ra[:FC_SHOW])
                        kn = " / ".join(cols[i] for i in ki2)
                        for k in sorted(want_k):
                            ra_ = [r for r in ga
                                   if tuple(r[i] for i in ki2) == k]
                            rb_ = [r for r in gb
                                   if tuple(r[i] for i in ki2) == k]
                            tag = " / ".join(fc_mask(cols[i], v, not mask)
                                             for i, v in zip(ki2, k))
                            out.append((ok, pad + f"{kn} = {tag}  —— "
                                        f"ita {len(ra_)} 列, "
                                        f"client {len(rb_)} 列"))
                            # 組內每一欄都一樣的就不必列, 只留有變化的
                            vary = [c for j, c in enumerate(cols)
                                    if len({r[j] for r in ra_ + rb_}) > 1]
                            if not vary:
                                out.append((ok, pad + "   兩邊這幾列的每一欄"
                                            "都一樣, 只差列數"))
                                continue
                            wv = [max(len(c), 12) for c in vary]
                            out.append((ok, pad + "   " + "     ".join(
                                f"{c:<{w_}}" for c, w_ in zip(vary, wv))))
                            for who, grp in (("ita", ra_), ("client", rb_)):
                                for r in grp[:20]:
                                    flag = ("   << 對方沒有這一列"
                                            if who == "ita"
                                            and miss.get(tuple(r)) else "")
                                    out.append((ok, pad + f"   {who:<6} "
                                                + "  ".join(
                                                    f"{r[cols.index(c)][:w_]:<{w_}}"
                                                    for c, w_ in zip(vary, wv))
                                                + flag))
                if not hint and how == "同期" and only_a == na and na:
                    out.append((ok, pad + "完全沒有交集, 而且湊不出對應的列 —— "
                                          "先確認配對是不是挑錯檔"))
                # 月份覆蓋只在「差很多」的時候才有意義; 差一兩列時列出全年
                # 十二個月只是噪音。
                if how != "同期" and ha and only_a > max(50, na * 0.05):
                    pick, cnt = fc_periods(b, cols)
                    if cnt:
                        want_ = sorted(a["span"])[0]
                        got = sorted(cnt)
                        out.append((ok, pad + f"對方那份的 {pick} 涵蓋 "
                                    + ", ".join(f"{y}-{m:02d}({cnt[(y, m)]:,})"
                                                for y, m in got[:8])
                                    + (" …" if len(got) > 8 else "")
                                    + ("" if want_ in cnt else
                                       f"   !! 沒有 {want_[0]}-{want_[1]:02d}")))
            if not ok:
                cap = max(0, FC_DIFF_CAP - len(diffs))
                sides = [("ita", a, ha)]
                if how == "同期":
                    sides.append(("client", b, hb))
                for src, u_, hs in sides:
                    if hs and cap:
                        for fn, row in fc_examples(
                                u_, cols, hs, min(cap, 200), lo,
                                dg if src == "client" else dga):
                            diffs.append(dict(row, 階段=st, 期間=fc_label(a),
                                               哪一邊=src, 檔名=fn))
                        cap = max(0, FC_DIFF_CAP - len(diffs))
    bar.close()
    add("")
    add("  內容比對 (整列取雜湊, 列的順序不影響)"
        + ("   --loose: 多值欄位先排序" if loose else ""))
    if loose_hit:
        add(f"    先排序過的欄位 ({len(loose_hit)}): "
            + ", ".join(fc_cname(c) for c in sorted(loose_hit)))
    for lab, dg in deg_hit.items():
        add(f"    [{lab}] 降階: 把 client 那一邊降到 ita 檔的精度再比")
        for c, (why, _, both) in sorted(dg.items()):
            add(f"       {fc_cname(c)}  {why}"
                + ("   [兩側都套]" if both else "   [只套 client]"))
        add("       !! 這樣只能證明「在 ita 檔還保留的精度上沒有差異」, "
            "不等於逐字一致")
    for ok, line in out:
        add(line)
    if diffs:
        d = root / FC_DIFF
        head = ["階段", "期間", "哪一邊", "檔名"]
        for r in diffs:                      # 不同組的欄位不見得一樣, 取聯集
            head += [k for k in r if k not in head]
        with open(d, "w", newline="", encoding="utf-8-sig") as f:
            w_ = csv.writer(f)
            w_.writerow(head)
            w_.writerows([[r.get(k, "") for k in head] for r in diffs])
        add("")
        add(f"  對不上的列寫到 {FC_DIFF} ({len(diffs):,} 列"
            + (", 已達上限" if len(diffs) >= FC_DIFF_CAP else "") + ")")
        add("  !! 裡面有票號, 只走內部管道")
    bar.close()
    add("")
    add(f"  {'全部一致' if not bad else f'!! {bad} 組對不上'}")


# -------------------------------------------------------------- --lists
# 對方是按幾份檔案清單分輪跑的。要照同樣的切法重跑, 得先確認每一份清單點
# 名的檔我們都有, 再按清單各自估一輪要多少磁碟 —— 一次全灌放不下, 分輪跑
# 才有機會。
def _stem(n: str) -> str:
    """檔名正規化成比對用的鍵: 去路徑、去副檔名、空白收斂、轉小寫。

    清單裡寫的是 .csv, 我們手上是 .parquet; 有些還帶了路徑或引號。不統一
    就會全部對不上, 而且看起來像「清單的檔一個都沒有」。
    """
    n = str(n).strip().strip('"').replace("\\", "/").split("/")[-1]
    for ext in (".csv", ".parquet", ".xlsx", ".xls", ".txt"):
        if n.lower().endswith(ext):
            n = n[: -len(ext)]
            break
    return " ".join(n.split()).lower()


def fl_find(root: Path) -> list:
    """找清單檔。檔名底線或空白都有 ("KA filelist.csv"), 一律不挑。"""
    out = []
    for d in (root, root / MISC, root / DATA_DIR):
        if d.is_dir():
            out += [p for p in d.glob("*.csv")
                    if "filelist" in p.name.lower()
                    .replace(" ", "").replace("_", "").replace("-", "")]
    return sorted(set(out), key=lambda p: p.name.lower())


def fl_read(p: Path, have: dict):
    """回 (檔名list, 用了第幾欄)。挑跟 data/ 對得上最多的那一欄 —— 清單
    可能有好幾欄 (序號、日期、備註), 猜錯欄就整份對不上。"""
    with open(p, newline="", encoding="utf-8-sig", errors="replace") as f:
        rows = list(csv.reader(f))
    if not rows:
        return [], 0
    ncol = max(len(r) for r in rows)
    best, bi = -1, 0
    for i in range(ncol):
        hit = sum(1 for r in rows if i < len(r) and _stem(r[i]) in have)
        if hit > best:
            best, bi = hit, i
    seen, out = set(), []
    for r in rows:
        if bi >= len(r):
            continue
        v = r[bi].strip()
        if not v or _stem(v) in seen:
            continue
        seen.add(_stem(v))
        out.append(v)
    # 第一列如果是表頭就丟掉。原本只在「底下對得上很多」時才丟, 但整份清
    # 單的檔都還沒進 data/ 時一個都對不上, 表頭就會被當成一個檔。改成看它
    # 長得像不像檔名 (沒有副檔名也沒有數字 = 表頭)。
    if out and _stem(out[0]) not in have:
        h = out[0].lower()
        if not any(h.endswith(e) for e in (".csv", ".parquet", ".xlsx", ".txt")) \
                and not any(ch.isdigit() for ch in h):
            out = out[1:]
    # 清單常常是用 dir 產出來的, 會把清單檔自己也列進去
    out = [n for n in out if "filelist" not in _stem(n).replace(" ", "")]
    return out, bi


def fl_size(paths: list):
    """回 (列數, 純文字位元組, 欄數)。逐檔量, 不取樣 —— 估磁碟就是要準。"""
    import pyarrow as pa
    import pyarrow.parquet as pq
    import pyarrow.compute as pc
    nrow = txt = ncol = 0
    bar = tqdm(paths, desc="  量", unit="檔", leave=False)
    for p in bar:
        if p.suffix.lower() != ".parquet":
            nrow_p = 0
            try:
                txt += p.stat().st_size
            except OSError:
                pass
            continue
        h = pq.ParquetFile(p)
        n = h.metadata.num_rows
        nrow += n
        if not n:
            continue
        b = next(h.iter_batches(batch_size=20000), None)
        if b is None:
            continue
        t = pa.Table.from_batches([b])
        ncol = max(ncol, t.num_columns)
        per = 0
        for c in t.columns:
            try:
                per += pc.sum(pc.binary_length(c)).as_py() or 0
            except Exception:
                per += c.nbytes
        txt += int(n * (per / t.num_rows))
    bar.close()
    return nrow, txt, ncol


def fl_nidx(root: Path):
    """回 (索引總數, 建在主表上的, 有沒有 DROP INDEX)。

    索引在 sqlite 裡不是小數目 —— 每筆約 30 bytes, 八千萬列乘一個索引就是
    2.4GB。但把全部索引加起來當峰值會高估: 腳本會 DROP, 而且大半索引建在
    衍生出來的小表上, 不是那張幾千萬列的主表。
    """
    ps = _patch_find(root, COMBINE_HINT)
    if not ps:
        return 3, 3, False
    t = ps[0].read_text(encoding="utf-8", errors="replace")
    idx = re.findall(r"create\s+(?:unique\s+)?index\s+"
                     r"(?:if\s+not\s+exists\s+)?[\"\'`\[]?\w+[\"\'`\]]?"
                     r"\s+on\s+[\"\'`\[]?(\w+)", t, re.I)
    drop = bool(re.search(r"drop\s+index", t, re.I))
    # 主表 = 直接被灌入來源資料的那張。腳本裡叫什麼不一定, 但它一定是索引
    # 數最多的那幾張之一, 而且名字裡通常有 combin。找不到就取最多的那張。
    main = [c for c in idx if "combin" in c.lower()]
    per: Counter = Counter(c.lower() for c in idx)
    n_main = len(main) or (per.most_common(1)[0][1] if per else 0)
    return len(idx), n_main, drop


# 這兩個係數是從一次實跑量出來的, 不是推算的。2026-09-17 跑 KA 那一輪:
# 純文字 5.52GB / 29,888,500 列, sqlite 載入完 12.8GB, 建完重複偵測的表
# 15.6GB, 最後 15.7GB。
#   載入   12.8 / 5.52       = 2.32 倍純文字
#   重複結構 2.8GB / 29.89M  = 94 bytes/列
# 之前用「每欄 1.1 bytes 標頭」推的結果是 10.89GB, 低估四成 —— 腳本會加一
# 批衍生欄 (TKT_NUM_KEY、_FLIGHT_DATE_SORT、Rate_Range 之類), 那些在原始
# 資料裡看不到, 只能實測。
SQL_LOAD_X = 2.32          # sqlite 載入後 / 純文字
SQL_DUP_PER_ROW = 94       # 重複偵測那幾張表, 每列


def fl_sqlite(nrow: int, txt: int, ncol: int, nidx: int = 3) -> int:
    return int(txt * SQL_LOAD_X + nrow * SQL_DUP_PER_ROW)


def fl_idxsize(nrow: int, nidx: int) -> int:
    return 0                # 索引已經含在上面的實測倍數裡了


def mode_lists(root: Path) -> None:
    try:
        import pyarrow  # noqa: F401
    except ImportError:
        sys.exit("要先裝 pyarrow:  pip install pyarrow")
    d = root / DATA_DIR
    if not d.is_dir():
        sys.exit(f"找不到 {d}")
    # 連子資料夾一起掃 —— --split 之後檔案都在 data/cx, data/ka, data/ara
    # 底下, 只掃第一層會把每一個都報成「缺」。
    have: dict = {}
    for p in sorted(d.rglob("*.parquet")) + sorted(d.rglob("*.csv")):
        have.setdefault(_stem(p.name), p)

    ls = fl_find(root)
    add("")
    add("=" * 72)
    add(f"檔案清單對照   {len(ls)} 份清單  vs  {DATA_DIR}/")
    add("=" * 72)
    if not ls:
        add("  找不到任何 *filelist*.csv (root / misc / data 都掃過了)")
        return
    sub = sorted({q.parent.name for q in have.values() if q.parent != d})
    add(f"  {DATA_DIR}/ 共 {len(have)} 個檔"
        + (f"  (含子資料夾 {', '.join(sub)})" if sub else ""))
    add("")

    got: dict = {}
    claimed: set = set()
    for p in ls:
        names, col = fl_read(p, have)
        hit = [n for n in names if _stem(n) in have]
        miss = [n for n in names if _stem(n) not in have]
        got[p.name] = [have[_stem(n)] for n in hit]
        claimed |= {_stem(n) for n in hit}
        add(f"  {p.name}   (用第 {col + 1} 欄)   列了 {len(names)} 個檔")
        add(f"     {DATA_DIR}/ 有 {len(hit)}"
            + (f",  !! 缺 {len(miss)}" if miss else ",  全部都在"))
        for n in miss[:12]:
            add(f"       缺: {n[:64]}")
        if len(miss) > 12:
            add(f"       … 另外 {len(miss) - 12} 個")

    add("")
    add(f"  三份清單去重後共點名 {len(claimed)} 個檔")
    orphan = [v for k, v in have.items() if k not in claimed]
    if orphan:
        add(f"  !! {DATA_DIR}/ 有 {len(orphan)} 個檔沒出現在任何清單裡:")
        for q in orphan[:12]:
            add(f"       {q.name[:64]}")
        if len(orphan) > 12:
            add(f"       … 另外 {len(orphan) - 12} 個")
    else:
        add(f"  {DATA_DIR}/ 每個檔都被某份清單點到")

    # ---- 按清單估磁碟
    add("")
    # 先量完再印 —— 進度條跟表格同時往畫面寫, 表格會被切得七零八落
    n_all, n_main, has_drop = fl_nidx(root)
    rowsz, worst = [], 0
    for name, paths in got.items():
        nrow, txt, ncol = fl_size(paths)
        lo = fl_sqlite(nrow, txt, ncol)
        hi = lo + fl_idxsize(nrow, n_all)
        mid = lo + fl_idxsize(nrow, n_main)
        worst = max(worst, lo)
        rowsz.append((name, len(paths), nrow, txt, lo, mid, hi))
    add(f"  照這樣分輪跑, 每一輪的規模")
    add(f"    (sqlite 峰值 = 純文字 x {SQL_LOAD_X} + 每列 {SQL_DUP_PER_ROW}"
        " bytes, 係數取自 2026-09-17 KA 那一輪的實測)")
    add(f"    {'清單':<26}{'檔數':>6}{'列數':>15}{'純文字':>11}"
        f"{'sqlite 峰值':>13}")
    add("    " + "-" * 72)
    for name, nf, nrow, txt, lo, mid, hi in rowsz:
        add(f"    {name[:26]:<26}{nf:>6}{nrow:>15,}{mb(txt):>11}{mb(lo):>13}")
    add("    " + "-" * 72)
    try:
        import shutil
        free = shutil.disk_usage(str(root)).free
        add(f"    {'工作磁碟現在可用':<26}{'':>6}{'':>15}{'':>12}{mb(free):>14}")
        add("")
        if worst > free:
            add(f"  !! 最大的那一輪估 {mb(worst)}, 放不下 —— 還差 "
                f"{mb(worst - free)}")
        else:
            add(f"  最大的那一輪估 {mb(worst)}, 放得下 (剩 {mb(free - worst)})")
        add("  (不含產出的 Excel。跑完一輪要先把 sqlite 與產出收走再跑下一輪)")
    except Exception:
        pass
    add("  (估的是 sqlite 檔本身。腳本還會建暫存表, 實際峰值會再高一些;")
    add("   跑完一輪要先把 sqlite 刪掉再跑下一輪。)")


# -------------------------------------------------------------- --sample
# 換引擎的時候, 每撞到一個 DuckDB 不吃的寫法就要重跑一次, 光載入就 33 分
# 鐘。把 data/ 第一層的 parquet 換成各取前 N 列的小檔, 整條流程一兩分鐘就
# 跑完, 一輪把所有問題找齊。原檔搬到 data/_full/ (腳本只掃第一層, 看不到),
# 不是複製也不是刪除, --sample --restore 搬回來。
#
# 注意: 這是拿來驗「跑不跑得完」的, 不是驗結果。結果要靠完整跑完之後跟
# output_ka/ 逐格比對。
FULL_DIR = "_full"


def sm_files(d: Path) -> list:
    return sorted(p for p in d.glob("*.parquet") if p.is_file())


def mode_sample(root: Path, n: int, restore: bool) -> None:
    import pyarrow.parquet as pq

    d = root / DATA_DIR
    full = d / FULL_DIR
    if not d.is_dir():
        sys.exit(f"找不到 {d}")

    add("")
    add("=" * 72)
    add("抽樣模式   " + ("把原檔搬回來" if restore else f"每個檔取前 {n:,} 列"))
    add("=" * 72)

    if restore:
        if not full.is_dir():
            add("  data/_full/ 不在, 沒有東西要還原。")
            add("  (原檔本來就沒被搬走, 或者已經還原過了)")
            return
        back = sm_files(full)
        for p in sm_files(d):                  # 先清掉抽樣出來的小檔
            if (full / p.name).exists():
                p.unlink()
        for p in back:
            p.replace(d / p.name)
        try:
            full.rmdir()
        except OSError:
            add(f"  ! {full} 還有東西, 沒有刪掉, 自己看一下")
        add(f"  還原 {len(back)} 個檔。data/ 現在是完整資料。")
        return

    if full.is_dir() and sm_files(full):
        sys.exit(f"{full} 裡已經有檔 —— 上一次抽樣還沒還原。\n"
                 "  先跑 python csvtools.py --sample --restore")
    src = sm_files(d)
    if not src:
        sys.exit(f"{d} 第一層沒有 parquet")

    import pyarrow as pa

    # 上一次中途失敗留下的暫存檔先清掉, 不然會混在 data/ 裡被腳本讀到
    for stale in d.glob("*.sample"):
        stale.unlink()

    full.mkdir(exist_ok=True)
    done = rows_in = rows_out = 0
    for p in src:
        # Windows 不准改名一個還開著的檔。ParquetFile 跟 iter_batches 的
        # generator 都要先關掉, 不然 p.replace() 會丟 WinError 32。
        # (macOS/Linux 允許, 所以這個在本機測不出來)
        f = pq.ParquetFile(p)
        it = None
        try:
            total = f.metadata.num_rows
            it = f.iter_batches(batch_size=min(n, max(total, 1)))
            batch = next(it, None)
            tbl = (pa.Table.from_batches([batch]).slice(0, n)
                   if batch is not None else None)
        finally:
            if it is not None:
                it.close()
            f.close()
        del f, it

        if tbl is None:
            add(f"  ! {p.name} 是空的, 原樣搬走不抽樣")
            p.replace(full / p.name)
            continue
        tmp = p.with_name(p.name + ".sample")
        pq.write_table(tbl, tmp)
        try:
            p.replace(full / p.name)           # 原檔先收好, 再放小檔
        except OSError:
            tmp.unlink(missing_ok=True)        # 搬不動就別留半成品
            raise
        tmp.replace(p)
        done += 1
        rows_in += total
        rows_out += tbl.num_rows
    add(f"  {done} 個檔: {rows_in:,} 列 -> {rows_out:,} 列 "
        f"({rows_out / max(rows_in, 1):.2%})")
    add(f"  原檔在 {full}  (腳本只掃第一層, 看不到)")
    add("")
    add("  FX rate.csv 與 Station Region Mapping.csv 沒有動 —— 那是對照表,")
    add("  少了會讓腳本走到別的分支。")
    add("")
    add("  現在跑客戶腳本, 一兩分鐘就會跑完或是報錯:")
    add("      python combine_v2.35_w4Flight-2.py")
    add("  跑完 (不管成功失敗) 一定要還原:")
    add("      python csvtools.py --sample --restore")
    add("")
    add("  ! 這一輪的 output/ 是抽樣結果, 不能當成果, 還原後要重跑。")


# -------------------------------------------------------------- --split
# 三份清單互不重疊, 但三支腳本都只掃 data/ 的第一層。所以把各輪的檔放進
# data/ 底下的子資料夾 (腳本看不到), 要跑哪一輪就把那一輪搬上第一層。
# 同一個磁碟內搬檔是改目錄項目, 不複製內容, 不佔額外空間也幾乎不花時間。
# _full 是抽樣時把原檔收起來的地方。--split 用 rglob 掃, 不排除的話會把
# 收好的原檔當成待分輪的檔搬走, 抽樣就再也還原不回去了。
SKIP_DIRS = {"filechecks", "output", ".git", "_full"}


def sp_tag(name: str) -> str:
    n = Path(name).stem.lower().replace("filelist", " ").replace("file list", " ")
    return "".join(re.sub(r"[^a-z0-9]+", " ", n).split()) or "list"


def sp_scan(root: Path) -> dict:
    """root 底下每個 csv/parquet 現在在哪。同名時 parquet 優先。

    filechecks/ 一定要跳過 —— 那裡面是比對用的副本, 檔名跟 data/ 一模一
    樣, 掃進來會把錯的那一份搬走。
    """
    out: dict = {}
    for p in root.rglob("*"):
        if not p.is_file() or p.suffix.lower() not in (".parquet", ".csv"):
            continue
        if {x.lower() for x in p.relative_to(root).parts[:-1]} & SKIP_DIRS:
            continue
        k = _stem(p.name)
        old = out.get(k)
        if old is None or (old.suffix.lower() == ".csv"
                           and p.suffix.lower() == ".parquet"):
            out[k] = p
    return out


def sp_move(src: Path, dst: Path, dry: bool) -> None:
    if dry:
        return
    dst.parent.mkdir(parents=True, exist_ok=True)
    src.replace(dst)          # 同一個磁碟內是改目錄項目, 不搬資料


def mode_split(root: Path, dry: bool, use: str | None) -> None:
    d = root / DATA_DIR
    if not d.is_dir():
        sys.exit(f"找不到 {d}")
    ls = fl_find(root)
    if not ls:
        sys.exit("找不到任何 *filelist*.csv")
    where = sp_scan(root)

    tags: dict = {}
    for p in ls:
        names, _ = fl_read(p, where)
        tags[sp_tag(p.name)] = {_stem(n) for n in names}
    owners: Counter = Counter()
    for st in tags.values():
        owners.update(st)
    # 被兩份以上清單點到的 (lookup 那兩個) 永遠留在 data/ 第一層 —— 每一輪
    # 都要用, 搬來搬去沒意義, 放子資料夾還得複製兩份。
    shared = {k for k, v in owners.items() if v > 1}

    add("")
    add("=" * 72)
    add(("[試算] " if dry else "") + "按清單分輪   "
        + (f"啟用 {use}" if use else "把各輪的檔搬進 data/ 的子資料夾"))
    add("=" * 72)
    ghost = sorted(k for k in owners if k not in where)
    for t, st in tags.items():
        own = len(st - shared - set(ghost))
        add(f"  {t:<6} 清單 {len(st):>4} 項"
            + f"   -> 自己的檔 {own}"
            + (f",  共用 {len(st & shared & where.keys())}" if st & shared else "")
            + (f",  找不到 {len(st & set(ghost))}" if st & set(ghost) else ""))
    add(f"  共用檔留在 {DATA_DIR}/ 第一層: "
        + ", ".join(sorted(where[k].name for k in shared if k in where)))
    if ghost:
        # 清單點名了但整個 root 底下都找不到的。多半是表頭殘留或 Excel 匯
        # 出之類的非檔案項目, 不會被搬 —— 但要講出來, 不然「共用 3 個卻只
        # 列出 2 個名字」這種對不上的數字會讓人懷疑整個結果。
        add(f"  !! 清單裡有 {len(ghost)} 項在整個資料夾底下都找不到 "
            "(不會搬, 多半是表頭或非檔案項目):")
        for k in ghost[:10]:
            add(f"       {k[:64]}")
    add("")

    # 現在第一層擺的是哪一輪
    now = Counter()
    for q in d.glob("*"):
        if q.is_file() and q.suffix.lower() in (".parquet", ".csv"):
            k = _stem(q.name)
            if k in shared:
                continue
            for t, st in tags.items():
                if k in st:
                    now[t] += 1

    # 不帶 --use 就只報現況。原本這裡預設是「把所有檔收進子資料夾」, 但那
    # 會把剛啟用好的一輪又收回去 —— 跟 --check 之類的旗標寫在同一行就中招
    # 了。搬檔這種事要明講才做。
    if use is None:
        add("  現在 data/ 第一層:")
        if now:
            for t, n in now.most_common():
                add(f"    {t}  {n} 個檔  <- 現在啟用的是這一輪")
        else:
            add("    沒有資料檔 (三輪都收在子資料夾裡)")
        for t in tags:
            sub = d / t
            n = len(list(sub.glob("*"))) if sub.is_dir() else 0
            add(f"    {DATA_DIR}/{t}/  {n} 個檔")
        add("")
        add(f"  要啟用某一輪:  python csvtools.py --split --use <{'/'.join(tags)}>")
        add("  要全部收起來:  python csvtools.py --split --use none")
        return

    moved = miss = 0
    if use.lower() not in ("none", "stow"):
        if use not in tags:
            sys.exit(f"沒有 {use} 這一輪, 只有: {', '.join(tags)}, "
                     "或 none (全部收起來)")
        # 先把第一層不屬於這一輪的收回各自的子資料夾
        for p in sorted(d.glob("*")):
            if not p.is_file() or p.suffix.lower() not in (".parquet", ".csv"):
                continue
            k = _stem(p.name)
            if k in shared or k in tags[use]:
                continue
            own = [t for t, st in tags.items() if k in st]
            if not own:
                add(f"  -- {p.name[:56]}  不在任何清單裡, 原地不動")
                continue
            add(f"  收   {p.name[:52]}  ->  {own[0]}/")
            sp_move(p, d / own[0] / p.name, dry)
            moved += 1
        # 再把這一輪的搬上第一層
        for p in sorted((d / use).glob("*")) if (d / use).is_dir() else []:
            if not p.is_file():
                continue
            add(f"  上   {use}/{p.name[:48]}  ->  {DATA_DIR}/")
            sp_move(p, d / p.name, dry)
            moved += 1
    else:                       # --use none: 全部收進子資料夾
        for t, st in tags.items():
            for k in sorted(st):
                if k in shared:
                    continue
                p = where.get(k)
                if p is None:
                    add(f"  !! {t}: 找不到 {k[:56]}")
                    miss += 1
                    continue
                if p.parent == d / t:
                    continue                       # 已經在位子上
                if p.suffix.lower() == ".csv":
                    # 只有 csv 沒有 parquet 的, 多半是轉換失敗的那幾個
                    add(f"  !! {t}: {p.name[:52]} 只有 CSV 沒有 parquet, "
                        "先不搬")
                    miss += 1
                    continue
                add(f"  搬   {p.name[:52]}  ->  {DATA_DIR}/{t}/")
                sp_move(p, d / t / p.name, dry)
                moved += 1
        # 共用檔搬回第一層
        for k in sorted(shared):
            p = where.get(k)
            if p is not None and p.parent != d:
                add(f"  搬   {p.name[:52]}  ->  {DATA_DIR}/ (共用)")
                sp_move(p, d / p.name, dry)
                moved += 1

    add("")
    add(f"  {'會搬' if dry else '搬了'} {moved} 個檔"
        + (f",  {miss} 個有問題" if miss else ""))
    if dry:
        add("  這是試算, 什麼都沒動。確認沒問題就把 --dry 拿掉。")
    else:
        add(f"  要跑某一輪:  python csvtools.py --split --use <{'/'.join(tags)}>")


# ---------------------------------------------------------------- --sql
# 磁碟估算裡索引佔了一大半, 但索引不影響 SQL 的結果, 只影響速度 —— 那是
# 純效能產物, 動它不會動到 re-performance 的邏輯。所以先看清楚: 這些索引
# 建在哪張表上、建完有沒有丟掉。順便盤點 sqlite 專屬語法, 評估換引擎的
# 成本。
SQLITE_ONLY = [
    ("rowid", r"\browid\b", "DuckDB 沒有隱含 rowid, 用到就得改寫"),
    ("julianday", r"\bjulianday\s*\(", "DuckDB 沒這個函數"),
    ("PRAGMA", r"\bpragma\b", "DuckDB 不吃, 直接刪掉即可"),
    ("INSERT OR", r"\binsert\s+or\s+(replace|ignore|rollback|abort|fail)\b",
     "DuckDB 語法不同"),
    ("AUTOINCREMENT", r"\bautoincrement\b", "DuckDB 用 SEQUENCE"),
    ("sqlite_master", r"\bsqlite_master\b", "DuckDB 用 duckdb_tables()"),
    ("ATTACH", r"\battach\s+database\b", "語法不同"),
    ("GROUP_CONCAT", r"\bgroup_concat\s*\(", "DuckDB 有, 分隔符參數要確認"),
    ("strftime", r"\bstrftime\s*\(", "DuckDB 有, 格式字串要確認"),
    ("printf", r"\bprintf\s*\(", "DuckDB 叫 format()"),
    ("substr", r"\bsubstr\s*\(", "兩邊都有, 通常不用改"),
    ("IFNULL", r"\bifnull\s*\(", "兩邊都有"),
]


PY_SQLITE_API = [
    ("sqlite3.connect", r"sqlite3\.connect\s*\(", "換成 duckdb.connect"),
    ("pandas read_sql", r"read_sql(?:_query|_table)?\s*\(",
     "DuckDB 用 con.execute(sql).df()"),
    ("pandas to_sql", r"\.to_sql\s*\(",
     "DuckDB 用 con.register() + CREATE TABLE AS SELECT"),
    ("executemany", r"\.executemany\s*\(", "DuckDB 有, 通常照用"),
    ("executescript", r"\.executescript\s*\(", "DuckDB 沒有, 要拆成多句"),
    ("cursor()", r"\.cursor\s*\(\)", "DuckDB 有, 通常照用"),
    ("commit()", r"\.commit\s*\(\)", "DuckDB 自動提交, 留著無害"),
    ("execute(", r"\.execute\s*\(", "兩邊都有"),
    ("fetchall/one/many", r"\.fetch(?:all|one|many)\s*\(", "兩邊都有"),
]


def sq_grab(t: str, pat: str) -> list:
    return [(t[:m.start()].count("\n") + 1, m)
            for m in re.finditer(pat, t, re.I | re.S)]


NEED_SEE = [
    ("連線那一行", r"sqlite3\.connect\s*\([^)]*\)"),
    ("rowid 用在哪", r"\browid\b"),
    ("INSERT OR", r"\binsert\s+or\s+\w+\s+into\b"),
]


def sq_show(t: str, label: str, pat: str, ctx: int = 2, cap: int = 6) -> None:
    lines = t.splitlines()
    hits = sq_grab(t, pat)[:cap]
    add(f"     [{label}]  {len(sq_grab(t, pat))} 處")
    for ln, _m in hits:
        for i in range(max(1, ln - ctx), min(len(lines), ln + ctx) + 1):
            mark = ">>" if i == ln else "  "
            add(f"       {mark} {i:>5}  {lines[i - 1].rstrip()[:96]}")
        add("       " + "-" * 40)


def sq_numcmp(t: str) -> list:
    """SQL 裡「欄位 vs 純數字」的比較。

    我們的來源是 parquet 全文字, 所以每一欄在資料庫裡都是 TEXT。sqlite 的
    TEXT 跟數字比較永遠不成立 ('007' = 7 是 false), DuckDB 會隱式轉型變成
    true。這種寫法換引擎會改變結果。

    只能掃 SQL, 不能掃整個檔 —— 整個檔掃下去會把 timeout=120、nrows=5 這種
    Python 賦值全撈進來, 三十幾個全是假警報, 看了等於沒看。
    """
    out = []
    for m in re.finditer(r"(\"\"\"|\'\'\'|\"|\')(.*?)\1", t, re.S):
        body = m.group(2)
        if not re.search(r"\b(select|where|update|delete|having|join)\b",
                         body, re.I):
            continue
        base = t[:m.start()].count("\n") + 1
        for c in re.finditer(r"\b([A-Za-z_]\w{2,})\s*"
                             r"(=|<>|!=|<=|>=|<|>)\s*(-?\d+)(?!\s*[\w.])",
                             body):
            out.append((base + body[:c.start()].count("\n"), c.group(0).strip()))
    return out


def mode_sql(root: Path, show: bool = False) -> None:
    """把腳本裡的 SQL 挑出來看。"""
    hits = []
    for hint, what in ((COMBINE_HINT, "合併"), (ARA_HINT, "月度彙總"),
                       (ADJ_HINT, "案件彙總")):
        for q in _patch_find(root, hint):
            hits.append((what, q))
    add("")
    add("=" * 72)
    add("腳本裡的 SQL")
    add("=" * 72)
    if not hits:
        add("  找不到任何腳本")
        return

    for what, q in hits:
        t = q.read_text(encoding="utf-8", errors="replace")
        add("")
        add(f"  [{what}] {q.name[:56]}   {len(t.splitlines()):,} 行")
        if "sqlite3" not in t.lower():
            add("     沒有用 sqlite —— 磁碟估算裡的 sqlite 數字對這支不適用")
            continue

        tabs = sq_grab(t, r"create\s+(temp\w*\s+)?table\s+"
                          r"(?:if\s+not\s+exists\s+)?[\"'`\[]?(\w+)")
        idxs = sq_grab(t, r"create\s+(?:unique\s+)?index\s+"
                          r"(?:if\s+not\s+exists\s+)?[\"'`\[]?(\w+)[\"'`\]]?"
                          r"\s+on\s+[\"'`\[]?(\w+)[\"'`\]]?\s*\(([^)]*)\)")
        dropi = sq_grab(t, r"drop\s+index\s+(?:if\s+exists\s+)?[\"'`\[]?(\w+)")
        dropt = sq_grab(t, r"drop\s+table\s+(?:if\s+exists\s+)?[\"'`\[]?(\w+)")

        add(f"     建表 {len(tabs)}  (其中暫存表 "
            f"{sum(1 for _, m in tabs if m.group(1))})"
            f"   建索引 {len(idxs)}   DROP INDEX {len(dropi)}"
            f"   DROP TABLE {len(dropt)}")

        if idxs:
            per: dict = {}
            for ln, m in idxs:
                per.setdefault(m.group(2).lower(), []).append(
                    (ln, m.group(1), " ".join(m.group(3).split())))
            add("     索引建在哪張表:")
            for tb, rows in sorted(per.items(), key=lambda x: -len(x[1])):
                add(f"       {tb:<24} {len(rows)} 個")
                for ln, nm, cols in rows:
                    add(f"          第 {ln:>5} 行  {nm[:28]:<28} ({cols[:44]})")
            if not dropi:
                add("     !! 沒有任何 DROP INDEX —— 所有索引會一直留在檔案裡,"
                    " 峰值就是全部相加")

        add("     sqlite 專屬語法:")
        none, must = True, 0
        for name, pat, note in SQLITE_ONLY:
            n = len(sq_grab(t, pat))
            if n:
                none = False
                add(f"       {name:<14} {n:>4} 處   {note}")
                if name in ("rowid", "julianday", "INSERT OR",
                            "AUTOINCREMENT", "sqlite_master", "ATTACH"):
                    must += n
        if none:
            add("       (沒有找到)")
        add("     Python 層用到的 sqlite API:")
        for name, pat, note in PY_SQLITE_API:
            n = len(sq_grab(t, pat))
            if n:
                add(f"       {name:<18} {n:>4} 處   {note}")
                if name in ("sqlite3.connect", "pandas read_sql",
                            "pandas to_sql", "executescript"):
                    must += n
        npragma = len(sq_grab(t, r"\bpragma\b"))
        add(f"     -> 換 DuckDB 要動的行數: 一定要改 約 {must} 處, "
            f"另外 {npragma} 行 PRAGMA 刪掉即可")
        nc = sq_numcmp(t)
        if nc:
            add(f"     !! 有 {len(nc)} 處是「欄位 vs 純數字」的比較。來源是"
                "parquet 全文字, 欄位在 DB 裡是 TEXT;")
            add("        sqlite 的 TEXT 跟數字比較永遠不成立, DuckDB 會隱式"
                "轉型 —— 換引擎會改變結果, 要逐處看:")
            for ln, txt in nc[:12]:
                add(f"          第 {ln:>5} 行   {txt[:60]}")
            if len(nc) > 12:
                add(f"          … 另外 {len(nc) - 12} 處")
        if show:
            add("")
            add("     ---- 需要人看的幾行 ----")
            for label, pat in NEED_SEE:
                if sq_grab(t, pat):
                    sq_show(t, label, pat)


# ------------------------------------------------------------- --space
# 跑完一輪之後「還有什麼佔著空間」不該用猜的。除了明顯的產出, 腳本中途
# 斷掉還會留下 sqlite 的暫存檔 (-journal / -wal), 那種檔跟資料庫本身一樣
# 大, 而且不會自己消失。
LEFTOVER = ("*.sqlite", "*.sqlite-journal", "*.sqlite-wal", "*.sqlite-shm",
            "*.db", "*.db-journal", "*.db-wal", "*.tmp", "~$*")


def sp_size(d: Path) -> tuple:
    n = tot = 0
    for p in d.rglob("*"):
        try:
            if p.is_file():
                tot += p.stat().st_size
                n += 1
        except OSError:
            pass
    return n, tot


def mode_space(root: Path) -> None:
    add("")
    add("=" * 72)
    add(f"空間盤點   {root}")
    add("=" * 72)

    rows = []
    for p in sorted(root.iterdir()):
        try:
            if p.is_dir():
                n, t = sp_size(p)
                rows.append((t, f"{p.name}/", n))
            elif p.is_file():
                rows.append((p.stat().st_size, p.name, 1))
        except OSError:
            pass
    rows.sort(reverse=True)
    add("  專案資料夾裡的東西 (由大到小)")
    tot = 0
    for t, name, n in rows[:14]:
        tot += t
        add(f"    {mb(t):>10}   {name[:48]:<48} {n:>6} 個檔")
    if len(rows) > 14:
        rest = sum(t for t, _, _ in rows[14:])
        tot += rest
        add(f"    {mb(rest):>10}   (其餘 {len(rows) - 14} 項)")
    add(f"    {mb(tot):>10}   合計")

    # 跑到一半留下來的
    left = []
    for pat in LEFTOVER:
        for p in root.rglob(pat):
            try:
                if p.is_file():
                    left.append((p.stat().st_size, p))
            except OSError:
                pass
    add("")
    if left:
        left.sort(reverse=True)
        add(f"  !! 上一輪留下來的暫存檔 {len(left)} 個, "
            f"共 {mb(sum(t for t, _ in left))} —— 可以直接刪:")
        for t, p in left[:10]:
            add(f"       {mb(t):>10}   {p.relative_to(root)}")
    else:
        add("  沒有留下暫存檔 (sqlite / journal / wal 都清乾淨了)")

    # 每一輪還要多少
    d = root / DATA_DIR
    subs = [x for x in sorted(d.iterdir()) if x.is_dir()] if d.is_dir() else []
    if subs:
        add("")
        add("  各輪還沒跑的話要多少空間")
        add(f"    {'':<10}{'列數':>15}{'純文字':>11}{'sqlite 峰值':>13}")
        for x in subs:
            ps = sorted(x.glob("*.parquet")) + sorted(x.glob("*.csv"))
            if not ps:
                continue
            nrow, txt, ncol = fl_size(ps)
            add(f"    {x.name:<10}{nrow:>15,}{mb(txt):>11}"
                f"{mb(fl_sqlite(nrow, txt, ncol)):>13}")
    try:
        import shutil
        free = shutil.disk_usage(str(root)).free
        add("")
        add(f"  工作磁碟可用   {mb(free)}")
    except Exception:
        pass


def main() -> None:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", default=".")
    ap.add_argument("--scripts", action="store_true",
                    help="解析 misc/ 裡的 .txt (其實是 Python)")
    ap.add_argument("--dup", action="store_true",
                    help="兩個資料夾的檔是不是同一批")
    ap.add_argument("--parquet", action="store_true", help="轉 parquet")
    ap.add_argument("--coverage", action="store_true",
                    help="每個系列涵蓋哪些月份, 哪裡有洞 (秒級)")
    ap.add_argument("--verify", action="store_true",
                    help="核對 parquet 與原始 CSV (刪原檔之前跑)")
    ap.add_argument("--schema", action="store_true",
                    help="每個系列的欄位版本 (只讀表頭, 秒級)")
    ap.add_argument("--typed", action="store_true",
                    help="--parquet 時讓 pyarrow 自己推型別 (預設全部存成文字)")
    ap.add_argument("--dry", action="store_true", help="--parquet 時只估算不轉")
    ap.add_argument("--quick", action="store_true", help="--dup 時只比大小")
    ap.add_argument("--dirs", default=None, help="--dup 的兩個資料夾, 逗號分隔")
    ap.add_argument("--only", default=None,
                    help="--parquet / --cmp 時只處理檔名含這段字的")
    ap.add_argument("--patch", action="store_true",
                    help="改客戶腳本: 讓它們吃 parquet")
    ap.add_argument("--paths", action="store_true",
                    help="--patch 時同時把輸入改 data/、輸出改 output/")
    ap.add_argument("--revert", action="store_true",
                    help="--patch 的還原 (從 *.py.bak)")
    ap.add_argument("--check", action="store_true",
                    help="開跑前體檢: 腳本改好沒、data/ 裡有什麼")
    ap.add_argument("--fc", action="store_true",
                    help="filechecks/ 底下 ita vs client 逐檔比對")
    ap.add_argument("--rows", action="store_true",
                    help="--fc 時真的逐行比內容 (要讀完所有檔)")
    ap.add_argument("--loose", action="store_true",
                    help='--fc / --cmp 時把 "|" 串起來的多值當成無序集合')
    ap.add_argument("--mask", action="store_true",
                    help="--fc 印出對不上的列時, 票號換成編號 (要貼到外面才用)")
    ap.add_argument("--degrade", action="store_true",
                    help="--fc 時把完好的一邊降到受損那一邊的精度再比")
    ap.add_argument("--sqlite", action="store_true",
                    help="--patch 時只把引擎換回 sqlite, 路徑等修改保留")
    ap.add_argument("--duck", action="store_true",
                    help="--patch 時把引擎換成 DuckDB (只改 import 那一行)")
    ap.add_argument("--show", action="store_true",
                    help="--sql 時把需要人判斷的那幾行連上下文印出來")
    ap.add_argument("--space", action="store_true",
                    help="看什麼佔著空間、有沒有上一輪留下的暫存檔")
    ap.add_argument("--sql", action="store_true",
                    help="盤點腳本裡的 SQL: 建表、索引、sqlite 專屬語法")
    ap.add_argument("--cpn", action="store_true",
                    help="查 coupon level 重複 (同一個票號+票聯出現多次), "
                         "分 exact 與 business 兩種並給例子")
    ap.add_argument("--where", default="filechecks/client",
                    help="--cpn 要掃哪個資料夾 (預設 filechecks/client)")
    ap.add_argument("--cpnrows", action="store_true",
                    help="把 cpn_examples.xlsx 裡那些票的原始列全部撈出來, "
                         "輸出 cpn_check.xlsx 給客戶對")
    ap.add_argument("--files", type=int, default=10,
                    help="--cpn 掃前幾個檔 (預設 10, 給 0 代表全部)")
    ap.add_argument("--eg", type=int, default=5,
                    help="--cpn 每個檔每種各取幾個例子 (預設 5)")
    ap.add_argument("--flat", action="store_true",
                    help="--fc 時 filechecks/ 是平鋪的一批 ita 檔, "
                         "用內容跟 data/ 配對 (不看檔名)")
    ap.add_argument("--sample", action="store_true",
                    help="把 data/ 換成各取前 N 列的小檔, 一兩分鐘跑完整條"
                         "流程找語法問題; 配 --restore 搬回原檔")
    ap.add_argument("--n", type=int, default=5000,
                    help="--sample 每個檔取幾列 (預設 5000)")
    ap.add_argument("--restore", action="store_true",
                    help="--sample 時把原檔搬回 data/")
    ap.add_argument("--split", action="store_true",
                    help="看目前啟用哪一輪; 配 --use 才會搬檔")
    ap.add_argument("--use", default=None,
                    help="--split 時啟用這一輪 (cx/ka/ara), "
                         "或 none 把三輪全部收進子資料夾")
    ap.add_argument("--lists", action="store_true",
                    help="對方的檔案清單 vs data/, 並按清單估磁碟")
    ap.add_argument("--scope", action="store_true",
                    help="對方的產出各自涵蓋哪些來源檔")
    ap.add_argument("--cmp", action="store_true",
                    help="我方產出 vs 對方產出")
    ap.add_argument("--ours", default=OUT_DIR)
    ap.add_argument("--theirs", default=MISC)
    args = ap.parse_args()
    root = Path(args.root).resolve()
    if not (args.scripts or args.dup or args.parquet or args.coverage
            or args.schema or args.verify or args.patch or args.cmp
            or args.check or args.scope or args.fc or args.lists or args.split or args.sql or args.space or args.sample or args.cpn or args.cpnrows):
        ap.error("要指定 --scripts / --dup / --coverage / --schema / --parquet"
                 " / --verify / --patch / --check / --scope / --fc / --cmp")

    t0 = time.time()
    add(f"csvtools  BUILD {BUILD}")
    add(f"  腳本 {Path(__file__).resolve()}")
    add(f"  資料 {root}")
    if args.scripts:
        mode_scripts(root)
    if args.dup:
        mode_dup(root, [x.strip() for x in args.dirs.split(",")]
                 if args.dirs else None, args.quick)
    if args.coverage:
        mode_coverage(root)
    if args.schema:
        mode_schema(root)
    if args.parquet:
        mode_parquet(root, args.dry, args.only, args.typed)
    if args.verify:
        mode_verify(root, args.only)
    if args.patch:
        mode_patch(root, args.dry, args.paths, args.revert,
                   args.duck, args.sqlite)
    if args.check:
        mode_check(root)
    if args.cpnrows:
        mode_cpnrows(root, args.where)
    if args.cpn:
        mode_cpn(root, args.where, args.eg, args.quick, args.files)
    if args.fc and args.flat:
        mode_fc_flat(root, args.rows)
    elif args.fc:
        mode_fc(root, args.rows, args.loose, args.only,
                args.degrade, args.mask)
    if args.lists:
        mode_lists(root)
    if args.sample:
        mode_sample(root, args.n, args.restore)
    if args.split:
        mode_split(root, args.dry, args.use)
    if args.sql:
        mode_sql(root, args.show)
    if args.space:
        mode_space(root)
    if args.scope:
        mode_scope(root, args.theirs)
    if args.cmp:
        mode_cmp(root, args.ours, args.theirs, args.only, args.quick,
                 args.loose)
    add("")
    add(f"完成, 共 {time.time() - t0:.0f}s")
    o = root / REPORT
    o.write_text("\n".join(_lines), encoding="utf-8")
    print(f"\n報告: {o}")


if __name__ == "__main__":
    main()
