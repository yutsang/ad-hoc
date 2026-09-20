#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
shrink.py —— 把一個資料夾底下的 CSV / Excel 就地轉成 parquet, 核對無誤才刪原檔。

    python shrink.py <資料夾>              掃描 + 轉檔 + 核對 (不刪任何東西)
    python shrink.py <資料夾> --dry        只看會轉什麼, 一個檔都不碰
    python shrink.py <資料夾> --rm         核對通過的 CSV 原檔才刪
    python shrink.py <資料夾> --rm --rm-excel   連 Excel 原檔也刪 (見下)

parquet 寫在原檔旁邊, 同名不同副檔名。核對三件事: 列數、欄名、前 200 列
逐格比對。任何一項對不上就不算通過, 那個檔的原檔絕對不會被刪。

**CSV 與 Excel 不一樣**:
  CSV   本來就只有格子的值, 轉 parquet 不會少東西。
  Excel 公式、數字格式、合併儲存格、欄寬、多工作表的結構全部留不住,
        parquet 只留值。所以 --rm 不碰 Excel, 要刪得再加 --rm-excel,
        而且那等於把原始交付檔換成一份只有值的副本 —— 進稽核底稿的東西
        想清楚再刪。

全部存成文字。讓 pyarrow 自己推型別的話, 前導零會掉 ("00123" -> 123),
而且不同檔推出不同型別, 之後合併會炸。
"""
from __future__ import annotations

import argparse
import csv
import sys
import time
from pathlib import Path

try:
    from tqdm import tqdm
except ImportError:
    class _Null:
        def update(self, *a):
            pass

        def close(self):
            pass

    def tqdm(x=None, **k):
        return x if x is not None else _Null()

SAMPLE = 200          # 逐格比對前幾列
SKIP_DIRS = {".git", "__pycache__", "$RECYCLE.BIN", "System Volume Information"}
_lines: list[str] = []


def add(s: str = "") -> None:
    _lines.append(s)
    print(s, flush=True)


def mb(n: int) -> str:
    return f"{n / 1e6:,.1f}MB" if n < 1e9 else f"{n / 1e9:,.2f}GB"


def csv_header(p: Path) -> list[str]:
    """表頭原樣取出 —— 不要 strip。

    pyarrow 用的是原樣的欄名 (可能帶前後空白)。strip 掉再拿去當 column_types
    的鍵就對不上, 那一欄會退回型別推斷, 票號 "0000000000000" 直接變成 0,
    而且不會報錯。
    """
    with open(p, "rb") as f:
        line = f.readline().rstrip(b"\r\n").decode("utf-8", "replace")
    return next(csv.reader([line.lstrip("﻿")]))


def _cell(v) -> str:
    return "" if v is None else str(v)


def out_name(src: Path, sheet: str | None) -> Path:
    if sheet is None:
        return src.with_suffix(".parquet")
    safe = "".join(c if c.isalnum() or c in " -_" else "_" for c in sheet)
    return src.with_name(f"{src.stem}__{safe.strip()}.parquet")


def conv_csv(src: Path, dst: Path) -> int:
    import pyarrow as pa
    import pyarrow.csv as pacsv
    import pyarrow.parquet as pq
    co = pacsv.ConvertOptions(
        column_types={c: pa.string() for c in csv_header(src)})
    w = None
    n = 0
    try:
        for b in pacsv.open_csv(
                src, read_options=pacsv.ReadOptions(block_size=1 << 26),
                convert_options=co):
            t = pa.Table.from_batches([b])
            bad = [f.name for f in t.schema if f.type != pa.string()]
            if bad:
                raise ValueError("這幾欄沒被指定成文字, 會失真: "
                                 + ", ".join(bad[:6]))
            if w is None:
                w = pq.ParquetWriter(dst, t.schema, compression="zstd")
            w.write_table(t)
            n += t.num_rows
    finally:
        if w is not None:
            w.close()
    return n


def xl_sheets(src: Path) -> list:
    from openpyxl import load_workbook
    wb = load_workbook(src, read_only=True, data_only=True)
    try:
        return list(wb.sheetnames)
    finally:
        wb.close()


def conv_xl(src: Path, sheet: str, dst: Path) -> int:
    import pyarrow as pa
    import pyarrow.parquet as pq
    from openpyxl import load_workbook
    wb = load_workbook(src, read_only=True, data_only=True)
    try:
        it = wb[sheet].iter_rows(values_only=True)
        hdr = [_cell(v) for v in next(it, ())]
        if not hdr:
            return 0
        # 重複或空白的欄名 parquet 收不了, 補成唯一
        seen: dict = {}
        cols = []
        for i, h in enumerate(hdr):
            h = h.strip() or f"col{i + 1}"
            if h in seen:
                seen[h] += 1
                h = f"{h}_{seen[h]}"
            else:
                seen[h] = 0
            cols.append(h)
        buf = [[] for _ in cols]
        n = 0
        w = None
        for row in it:
            if row is None or all(v is None for v in row):
                continue
            for i in range(len(cols)):
                buf[i].append(_cell(row[i]) if i < len(row) else "")
            n += 1
            if n % 200000 == 0:
                t = pa.table({c: buf[i] for i, c in enumerate(cols)},
                             schema=pa.schema([(c, pa.string()) for c in cols]))
                if w is None:
                    w = pq.ParquetWriter(dst, t.schema, compression="zstd")
                w.write_table(t)
                buf = [[] for _ in cols]
        t = pa.table({c: buf[i] for i, c in enumerate(cols)},
                     schema=pa.schema([(c, pa.string()) for c in cols]))
        if w is None:
            w = pq.ParquetWriter(dst, t.schema, compression="zstd")
        w.write_table(t)
        w.close()
        return n
    finally:
        wb.close()


def check_csv(src: Path, dst: Path) -> list[str]:
    import pyarrow.parquet as pq
    bad = []
    pf = pq.ParquetFile(dst)
    pcols = list(pf.schema_arrow.names)
    if pcols != csv_header(src):
        bad.append(f"欄名不同 ({len(pcols)} vs {len(csv_header(src))} 欄)")
        return bad
    with open(src, "rb") as f:
        f.readline()
        want = []
        for _ in range(SAMPLE):
            ln = f.readline()
            if not ln:
                break
            want.append(next(csv.reader(
                [ln.rstrip(b"\r\n").decode("utf-8", "replace")])))
    b = next(pf.iter_batches(batch_size=SAMPLE), None)
    got = ([[_cell(v) for v in r]
            for r in zip(*[b.column(c).to_pylist() for c in pcols])]
           if b is not None else [])
    for i, (g, wnt) in enumerate(zip(got, want)):
        if g != wnt:
            d = [(pcols[k], wnt[k], g[k])
                 for k in range(min(len(g), len(wnt))) if g[k] != wnt[k]]
            bad.append(f"第 {i + 2} 列有 {len(d)} 格對不上"
                       + (f": {d[0][0]} 原={d[0][1]!r} parquet={d[0][2]!r}"
                          if d else ""))
            break
    return bad


def check_xl(src: Path, sheet: str, dst: Path) -> list[str]:
    import pyarrow.parquet as pq
    from openpyxl import load_workbook
    bad = []
    pf = pq.ParquetFile(dst)
    pcols = list(pf.schema_arrow.names)
    wb = load_workbook(src, read_only=True, data_only=True)
    try:
        it = wb[sheet].iter_rows(values_only=True)
        next(it, None)
        want, nrow = [], 0
        for row in it:
            if row is None or all(v is None for v in row):
                continue
            nrow += 1
            if len(want) < SAMPLE:
                want.append([_cell(row[i]) if i < len(row) else ""
                             for i in range(len(pcols))])
    finally:
        wb.close()
    if pf.metadata.num_rows != nrow:
        bad.append(f"[{sheet}] 列數不同  原檔 {nrow:,} / "
                   f"parquet {pf.metadata.num_rows:,}")
        return bad
    b = next(pf.iter_batches(batch_size=SAMPLE), None)
    got = ([list(r) for r in zip(*[b.column(c).to_pylist() for c in pcols])]
           if b is not None else [])
    for i, (g, wnt) in enumerate(zip(got, want)):
        if g != wnt:
            d = [(pcols[k], wnt[k], g[k])
                 for k in range(len(g)) if g[k] != wnt[k]]
            bad.append(f"[{sheet}] 第 {i + 2} 列有 {len(d)} 格對不上"
                       + (f": {d[0][0]} 原={d[0][1]!r} parquet={d[0][2]!r}"
                          if d else ""))
            break
    return bad


def main() -> None:
    ap = argparse.ArgumentParser(description="CSV/Excel 就地轉 parquet")
    ap.add_argument("path")
    ap.add_argument("--dry", action="store_true", help="只看會轉什麼")
    ap.add_argument("--rm", action="store_true",
                    help="核對通過的 CSV 原檔才刪")
    ap.add_argument("--rm-excel", action="store_true",
                    help="連 Excel 原檔也刪 (公式與格式會永久消失)")
    a = ap.parse_args()
    root = Path(a.path).resolve()
    if not root.is_dir():
        sys.exit(f"不是資料夾: {root}")
    try:
        import pyarrow  # noqa: F401
    except ImportError:
        sys.exit("要先裝 pyarrow:  pip install pyarrow")

    t0 = time.time()
    add(f"shrink   {root}")
    todo = []
    for p in sorted(root.rglob("*")):
        if not p.is_file() or p.name.startswith("~$"):
            continue
        if {x for x in p.relative_to(root).parts[:-1]} & SKIP_DIRS:
            continue
        if p.suffix.lower() in (".csv", ".xlsx", ".xlsm", ".xls"):
            todo.append(p)
    add("=" * 72)
    add(f"找到 {len(todo)} 個可以轉的檔"
        f"  (CSV {sum(1 for x in todo if x.suffix.lower() == '.csv')}"
        f"  Excel {sum(1 for x in todo if x.suffix.lower() != '.csv')})")
    add("=" * 72)
    if not todo:
        return
    if a.dry:
        for p in todo[:40]:
            add(f"  {mb(p.stat().st_size):>10}   {p.relative_to(root)}")
        if len(todo) > 40:
            add(f"  … 另外 {len(todo) - 40} 個")
        add("")
        add(f"  原檔共 {mb(sum(p.stat().st_size for p in todo))}")
        add("  這是試算, 什麼都沒動。")
        return

    ok, bad, skipped = [], [], []
    src_b = got_b = 0
    bar = tqdm(total=sum(p.stat().st_size for p in todo), unit="B",
               unit_scale=True, desc="  轉換", leave=False)
    for p in todo:
        sz = p.stat().st_size
        bar.update(sz)
        is_csv = p.suffix.lower() == ".csv"
        try:
            sheets = [None] if is_csv else xl_sheets(p)
            outs, probs, made = [], [], []
            for sh in sheets:
                dst = out_name(p, sh)
                # 已經有 parquet 就不重轉, 但一樣要核對過才算數。這樣「先轉
                # 檔、確認沒問題、再回頭刪」這個流程才走得通 —— 不然第二次
                # 跑會因為檔案已存在而全部跳過, --rm 一個都刪不掉。
                if not dst.exists():
                    (conv_csv(p, dst) if is_csv else conv_xl(p, sh, dst))
                    made.append(dst)
                probs += (check_csv(p, dst) if is_csv
                          else check_xl(p, sh, dst))
                outs.append(dst)
            if probs:
                for d in made:          # 只收掉這次自己寫出來的半成品
                    d.unlink(missing_ok=True)
                bad.append((p, probs))
            elif not outs:
                skipped.append(p)
            else:
                ok.append((p, outs, is_csv))
                src_b += sz
                got_b += sum(d.stat().st_size for d in outs)
        except Exception as e:
            bad.append((p, [f"{type(e).__name__}: {e}"]))
    bar.close()

    add("")
    add(f"  轉好 {len(ok)}   有問題 {len(bad)}   跳過 {len(skipped)}")
    if src_b:
        add(f"  原檔 {mb(src_b)}  ->  parquet {mb(got_b)}"
            f"   ({src_b / got_b:.1f}x, 刪掉原檔可省 {mb(src_b - got_b)})")
    for p, probs in bad[:15]:
        add(f"  !! {p.relative_to(root)}")
        for x in probs[:3]:
            add(f"       {x}")
    if len(bad) > 15:
        add(f"  … 另外 {len(bad) - 15} 個有問題")
    for p in skipped[:8]:
        add(f"  -- 跳過 (沒有可轉的工作表): {p.relative_to(root)}")

    if not a.rm:
        add("")
        add("  沒有刪任何東西。核對過的原檔要刪就加 --rm")
        add("  (Excel 的原檔要再加 --rm-excel —— 公式與格式會永久消失)")
        return

    freed = n_del = 0
    kept_xl = 0
    for p, outs, is_csv in ok:
        if not is_csv and not a.rm_excel:
            kept_xl += 1
            continue
        sz = p.stat().st_size
        p.unlink()
        freed += sz
        n_del += 1
    add("")
    add(f"  刪掉 {n_del} 個核對通過的原檔, 釋出 {mb(freed)}")
    if kept_xl:
        add(f"  {kept_xl} 個 Excel 原檔保留 (要刪加 --rm-excel)")
    if bad:
        add(f"  !! {len(bad)} 個沒通過的原檔一個都沒動")

    add("")
    add(f"完成, 共 {time.time() - t0:.0f}s")
    o = root / "shrink_report.txt"
    o.write_text("\n".join(_lines), encoding="utf-8")
    print(f"\n報告: {o}")


if __name__ == "__main__":
    main()
