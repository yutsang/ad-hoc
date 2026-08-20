#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
filecheck.py —— 兩批 CSV 交付檔的 row / column 比對 (audit check 用)。

情境: 同一批資料先後由兩個來源給過, 檔名規則不一樣, 同一段期間各自叫不同
名字, 而且舊來源裡混了好幾個系列。要先解析出期間、分開系列, 才談得上比內容。

放在專案主資料夾跑, 它自己會進 filechecks/ 底下那兩個子資料夾。
設定在下面那一段直接寫死 —— 這是一次性的檢查, 不做成可配置工具。

指令只有三個:
    python filecheck.py                 全套 (檔名/大小/SHA/欄位/列數, 有差才逐行)
    python filecheck.py --quick         只看檔名與大小 (秒級)
    python filecheck.py --only 2017Aug  只做檔名含這段字的

全部檔案都會掃過一次 (第四節「全檔清點」), 不是只掃配對到的那幾個 ——
底稿要的是「我們收到什麼」的完整紀錄。

輸出兩份:
    filecheck_report.txt     完整過程
    filecheck_summary.csv    每一組配對一列, 可以直接貼進底稿
    filecheck_inventory.csv  每個檔一列 (欄數/列數/SHA256), 含沒配對到的
    filecheck_diffs.csv      值不同的那幾列差在哪個欄 (⚠️ 含原始值)
兩份都只有檔名與統計值, 沒有任何一列資料內容。
"""
from __future__ import annotations

import argparse
import calendar
import csv
import hashlib
import re
import sys
import time
from collections import Counter
from pathlib import Path

class _NullBar:
    """小檔用這個 —— 真的 tqdm 沒給 total 會印一條沒用的 0it 條, 把報告沖花。"""

    def update(self, *a):
        pass

    def close(self):
        pass


try:
    from tqdm import tqdm
except ImportError:      # 沒裝就照跑, 只是沒有進度條
    def tqdm(*a, **k):
        return _NullBar()

# 我在 Mac 上寫、使用者在 Windows 上跑。檔案沒同步到就會跑到舊版而不自知,
# 而舊版的輸出看起來一樣像模像樣。所以每次跑都把版本印在最前面。
BUILD = "2026-09-15q  同月份全部併起來比"

# ======================================================= 設定 (一次性, 寫死)
ROOT_DIRNAME = "filechecks"
DIR_NEW = "client"        # 新來源的子資料夾名
DIR_OLD = "ita"           # 舊來源的子資料夾名
# 舊來源底下有三個系列, 只有這個跟新來源是同一批資料; 其餘的單獨列出不配對。
# 系列名 = 檔名去掉數字與月份之後剩下的字 (見 series_of), 這裡填其中一段即可。
OLD_SERIES = "version"
REPORT = "filecheck_report.txt"
SUMMARY = "filecheck_summary.csv"
DIFFS = "filecheck_diffs.csv"
INVENTORY = "filecheck_inventory.csv"
MAX_SHOW = 5              # 逐行比對時最多印幾個不同的行號
BAR_MIN = 20_000_000      # 檔案超過這個大小才顯示進度條
# ======================================================= 設定結束

MONTHS = {m: i for i, m in enumerate(
    ["jan", "feb", "mar", "apr", "may", "jun",
     "jul", "aug", "sep", "oct", "nov", "dec"], 1)}
MON = "jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec"
HALF = {1: "上半月", 2: "下半月", 0: "整月"}

_lines: list[str] = []
_rows: list[list] = []
_diffs: list[list] = []
_match: dict = {}   # 檔路徑 -> (對手檔名, 判定)


def add(s: str = "") -> None:
    _lines.append(s)
    print(s, flush=True)


def parse_name(name: str) -> tuple | None:
    """檔名 -> (年, 月, 半月)。半月 1=上半 2=下半 0=整月。認不出回 None。

    月份後面只吃字母不吃數字 —— 否則 "2017Aug1" 的尾碼 1 會被當成月份的
    一部分, 上下半月就判成整月了。
    """
    n = name.lower()
    if n.endswith(".csv"):
        n = n[:-4]
    m = re.search(r"[_\s](20\d{2})(0[1-9]|1[0-2])[_\s]", n)
    if m:
        return (int(m.group(1)), int(m.group(2)), 0)
    m = re.search(rf"(\d{{1,2}})\s*-\s*\d{{1,2}}\s*({MON})[a-z]*\s*(20\d{{2}})", n)
    if m:
        return (int(m.group(3)), MONTHS[m.group(2)],
                1 if int(m.group(1)) <= 1 else 2)
    n2 = re.sub(r"20\d{2}\s*version", " ", n)   # 開頭那個版本年份不是資料期間
    m = re.search(rf"(20\d{{2}})\s*({MON})[a-z]*\s*([12])?(?![0-9])", n2)
    if m:
        return (int(m.group(1)), MONTHS[m.group(2)],
                int(m.group(3)) if m.group(3) else 0)
    m = re.search(rf"({MON})[a-z]*\s*(20\d{{2}})", n2)
    if m:
        return (int(m.group(2)), MONTHS[m.group(1)], 0)
    return None


def series_of(name: str) -> str:
    """檔名去掉數字與月份之後剩下的字 = 系列名。

    順序要緊: 先把數字拿掉。"1-15aug2017" 這種月份與數字黏在一起的寫法,
    \\b 邊界在 "5" 與 "a" 之間不成立, 先拆月份會拆不掉, 結果同一個系列被
    按月拆成十幾個。
    """
    n = name.lower()
    if n.endswith(".csv"):
        n = n[:-4]
    n = re.sub(r"[\d_\-.]+", " ", n)
    n = re.sub(rf"\b({MON})[a-z]*\b", " ", n)
    return " ".join(n.split()) or "(無)"


# 找航班日期欄用的候選名稱, 依優先順序
DATECOLS = ["flight_date", "flight date", "flight_act_dep_date",
            "flight_sch_dep_date", "out_hkg_std", "flight yearmonth"]
_SCAN: dict = {}
_DAYS: dict = {}     # 檔 -> {日期: 列數}


def _field(line: bytes, i: int) -> bytes:
    """只取第 i 欄, 不做整行 split —— 28M 列 × 54 欄的 split 會慢很多,
    而日期欄通常排在前面, 掃到那裡就夠。"""
    pos = 0
    for _ in range(i):
        j = line.find(b",", pos)
        if j < 0:
            return b""
        pos = j + 1
    j = line.find(b",", pos)
    return line[pos:j if j >= 0 else len(line)]


def _day(v: bytes) -> str:
    """各種寫法收斂成 YYYY-MM-DD。認不出回空字串。"""
    t = v.strip().strip(b'"').decode("utf-8", "replace").strip()
    if not t:
        return ""
    t = t.replace("/", "-")
    if len(t) >= 10 and t[4] == "-" and t[7] == "-":
        return t[:10]
    if len(t) >= 10 and t[2] == "-" and t[5] == "-":      # DD-MM-YYYY
        return f"{t[6:10]}-{t[3:5]}-{t[0:2]}"
    if len(t) == 8 and t.isdigit():                        # YYYYMMDD
        return f"{t[:4]}-{t[4:6]}-{t[6:]}"
    if len(t) >= 7 and t[4] == "-":                        # 只有年月
        return t[:7]
    return ""


def scan(p: Path) -> tuple[str, bytes, int]:
    """一次讀完就同時拿 SHA256、表頭、資料列數、每日列數 —— 分幾次讀等於
    把檔案讀幾遍。二進位讀: 兩邊編碼未必一致, 解碼只會製造假差異。快取。"""
    if p in _SCAN:
        return _SCAN[p]
    h = hashlib.sha256()
    hdr, n, first = b"", 0, True
    idx = None
    days: Counter = Counter()
    # 小檔不開進度條 —— 一閃而過的條只會把報告文字沖花
    sz = p.stat().st_size
    bar = (tqdm(total=sz, unit="B", unit_scale=True,
                desc=f"  讀 {p.name[:38]}", leave=False)
           if sz >= BAR_MIN else _NullBar())
    with open(p, "rb") as f:
        for line in f:
            h.update(line)
            bar.update(len(line))
            if first:
                hdr, first = line.rstrip(b"\r\n"), False
                low = [c.strip().strip('"').lower()
                       for c in cols_of(hdr)]
                for want in DATECOLS:
                    if want in low:
                        idx = low.index(want)
                        break
            elif line.strip():
                n += 1
                if idx is not None:
                    d = _day(_field(line, idx))
                    if d:
                        days[d] += 1
    bar.close()
    _DAYS[p] = days
    _SCAN[p] = (h.hexdigest(), hdr, n)
    return _SCAN[p]


def cols_of(hdr: bytes) -> list[str]:
    return [c.strip().strip('"') for c in
            hdr.decode("utf-8", "replace").lstrip("﻿").split(",")]


def inorder_diff(a: Path, b: Path) -> tuple[int, list[int]]:
    """依序逐行比對, O(1) 記憶體。行尾換行先去掉 —— CRLF 與 LF 不是內容差異。"""
    ndiff, shown, ln = 0, [], 0
    with open(a, "rb") as fa, open(b, "rb") as fb:
        while True:
            la, lb = fa.readline(), fb.readline()
            if not la and not lb:
                break
            ln += 1
            if la.rstrip(b"\r\n") != lb.rstrip(b"\r\n"):
                ndiff += 1
                if len(shown) < MAX_SHOW:
                    shown.append(ln)
    return ndiff, shown


def rowset_diff(A: list[Path], B: list[Path]) -> tuple[int, int, int]:
    """不計順序的行集合差異 -> (只在 a, 只在 b, 兩邊都有)。

    用 Counter 不用 set: 同一行可能重複出現, set 會把「a 有三筆、b 有一筆」
    看成一樣。⚠️ 一個 160MB 的檔約百萬列, Counter 大約吃 150MB 記憶體。
    """
    def dg(line: bytes) -> bytes:
        return hashlib.blake2b(line.rstrip(b"\r\n"), digest_size=8).digest()

    c: Counter = Counter()
    for p in A:
        with open(p, "rb") as f:
            f.readline()
            for line in f:
                if line.strip():
                    c[dg(line)] += 1
    both = only_b = 0
    for p in B:
        with open(p, "rb") as f:
            f.readline()
            for line in f:
                if not line.strip():
                    continue
                d = dg(line)
                if c[d] > 0:
                    c[d] -= 1
                    both += 1
                else:
                    only_b += 1
    return sum(v for v in c.values() if v > 0), only_b, both


# 這些欄的值不印出來 —— 報告檔要能給人看, 不能帶個資
PII = {"pax_title", "pax_surname", "pax_given_name", "seat_number",
       "boarding_number", "booking_reference", "eb_emd_no", "endorsement",
       "matching key", "filter_pax", "tkt_num", "tkt no", "ticket number",
       "tkt_set_num", "orig_issue_tkt_num", "original orig issue tkt num",
       "coupon_number"}
DIFF_CAP = 3000      # 每邊最多留幾列獨有的行下來做欄位比對


def _fields(line: bytes) -> list[str]:
    try:
        return next(csv.reader([line.decode("utf-8", "replace")]))
    except Exception:
        return line.decode("utf-8", "replace").split(",")


def diff_detail(A: list[Path], B: list[Path]):
    """兩邊各自獨有的行配對起來, 看差在哪個欄。

    列數相同、檔案大小也相同, 卻有 N 列對不上 —— 那不是多了少了, 是同樣
    那幾列的值被改過。要講清楚「改了哪個欄、改成什麼」才有 audit 價值。

    做法: 先用行雜湊挑出兩邊各自獨有的行 (數量少, 留得下內容), 再兩兩配對,
    取欄位相同最多的那一個當同一列的前後版本。
    """
    def dg(line: bytes) -> bytes:
        return hashlib.blake2b(line.rstrip(b"\r\n"), digest_size=8).digest()

    c: Counter = Counter()
    for p in A:
        with open(p, "rb") as f:
            f.readline()
            for line in f:
                if line.strip():
                    c[dg(line)] += 1
    ob: list = []
    for p in B:
        with open(p, "rb") as f:
            f.readline()
            for line in f:
                if not line.strip():
                    continue
                d = dg(line)
                if c[d] > 0:
                    c[d] -= 1
                elif len(ob) < DIFF_CAP:
                    ob.append(line.rstrip(b"\r\n"))
    left = {k for k, v in c.items() if v > 0}
    oa: list = []
    for p in A:
        with open(p, "rb") as f:
            f.readline()
            for line in f:
                if line.strip() and dg(line) in left and len(oa) < DIFF_CAP:
                    oa.append(line.rstrip(b"\r\n"))

    fa = [_fields(x) for x in oa]
    fb = [_fields(x) for x in ob]
    used = set()
    pairs = []
    for i, ra in enumerate(fa):
        best, bj = -1, None
        for j, rb in enumerate(fb):
            if j in used or len(rb) != len(ra):
                continue
            k = sum(1 for x, y in zip(ra, rb) if x == y)
            if k > best:
                best, bj = k, j
        if bj is None:
            continue
        used.add(bj)
        d = [n for n, (x, y) in enumerate(zip(ra, fb[bj])) if x != y]
        pairs.append((ra, fb[bj], d))
    return len(oa), len(ob), pairs


KEYCOLS = [["tkt_num", "tkt no", "ticket number"],
           ["flight_date", "flight date"],
           ["flight_no", "flight number"]]


def _keyset(p: Path) -> set | None:
    """用 票號+航班日期+航班號 當鍵。欄位版本不同的兩個檔沒辦法比整行雜湊,
    但這三個欄哪個版本都有, 足以認出「同一筆記錄」。找不到欄就回 None。"""
    _, hdr, _ = scan(p)
    low = [c.strip().strip('"').lower() for c in cols_of(hdr)]
    idx = []
    for cands in KEYCOLS:
        j = next((low.index(c) for c in cands if c in low), None)
        if j is None:
            return None
        idx.append(j)
    out = set()
    with open(p, "rb") as f:
        f.readline()
        for line in f:
            if line.strip():
                out.add(tuple(_field(line, j).strip().strip(b'"') for j in idx))
    return out


def _overlap(a: Path, b: Path) -> bool:
    da, db = _DAYS.get(a) or {}, _DAYS.get(b) or {}
    return bool(da and db and (set(da) & set(db)))


def collect(base: Path, keep: str | None) -> tuple[dict, dict, list]:
    """回 ({(年,月): {半月: [路徑]}}, {系列: [路徑]}, 解析不出期間的)。
    keep 有值時只留系列名含那段字的檔。"""
    ok: dict = {}
    ser: dict = {}
    bad: list = []
    for p in sorted(base.rglob("*.csv"), key=lambda x: str(x).lower()):
        sn = series_of(p.name)
        ser.setdefault(sn, []).append(p)
        if keep and keep.lower() not in sn:
            continue
        k = parse_name(p.name)
        if k is None:
            bad.append(p)
        else:
            ok.setdefault(k[:2], {}).setdefault(k[2], []).append(p)
    return ok, ser, bad


def mb(n: int) -> str:
    return f"{n / 1e6:,.1f}MB"


def flat(d: dict) -> list[Path]:
    """{半月: [路徑]} 攤平成一個清單, 上半月在前。"""
    return [p for h in sorted(d) for p in d[h]]


def tail(name: str, w: int = 32) -> str:
    """這批檔名共用很長的前綴, 截頭會全部長一樣, 所以留尾巴。"""
    return name if len(name) <= w else "…" + name[-(w - 1):]


def _rec(ym, half, a, b, verdict, **kw) -> None:
    """一組配對寫一列到 summary.csv —— 底稿要的就是這張表。"""
    _rows.append([
        f"{ym[0]}-{ym[1]:02d}", HALF.get(half, "整月合併"),
        a.name if a else "", b.name if b else "",
        a.stat().st_size if a else "", b.stat().st_size if b else "",
        kw.get("sha", ""), kw.get("cn", ""), kw.get("co", ""),
        kw.get("csame", ""), kw.get("rn", ""), kw.get("ro", ""),
        kw.get("rsame", ""), kw.get("on", ""), kw.get("oo", ""), verdict,
    ])


def _cname(cols: list[str], i: int) -> str:
    """欄名可能是空的 (原檔就有沒命名的欄), 那就用位置。"""
    nm = cols[i].strip() if i < len(cols) else ""
    return nm or f"第{i + 1}欄(無欄名)"


def _same_multi(x: str, y: str) -> bool:
    """同一個欄裡用 | 串起來的多值, 兩邊只是排列順序不同 -> 內容其實一樣。

    實測就是這種: client='A|B' 而 ita='B|A'。行雜湊會不同, 於是兩邊各自
    都把這一列算成「只有我有」, 看起來像資料差異, 其實不是。
    """
    for sep in ("|", ";"):
        if sep in x or sep in y:
            return sorted(t.strip() for t in x.split(sep)) == \
                   sorted(t.strip() for t in y.split(sep))
    return False


def compare_file(a: Path, b: Path, quick: bool) -> tuple[str, list[str]]:
    """一個檔對一個檔。回 (判定, 要附註的細節行)。

    判定只有兩類: 一樣 / 不一樣。「一樣」底下再標原因 (逐位元 / 只有行序 /
    只有欄內多值順序) —— 那幾種都不是資料差異, 不必再往下追。
    """
    if quick:
        sa, sb = a.stat().st_size, b.stat().st_size
        return ("一樣 (大小相同)" if sa == sb
                else f"不一樣 (大小差 {abs(sa - sb):,} bytes)"), []

    ha, hda, na = scan(a)
    hb, hdb, nb = scan(b)
    ca, cb = cols_of(hda), cols_of(hdb)
    if ha == hb:
        return f"一樣 (逐位元相同)   {len(ca)} 欄 / {na:,} 列", []

    if ca != cb:
        d = []
        oa = [c for c in ca if c not in cb]
        ob = [c for c in cb if c not in ca]
        if oa:
            d.append(f"只有 {DIR_NEW} 有的欄 ({len(oa)}): {', '.join(oa[:15])}")
        if ob:
            d.append(f"只有 {DIR_OLD} 有的欄 ({len(ob)}): {', '.join(ob[:15])}")
        if not oa and not ob:
            mv = [(k, x, y) for k, (x, y) in enumerate(zip(ca, cb), 1) if x != y]
            d.append(f"欄名相同但順序不同, {len(mv)} 個位置")
        return f"不一樣 (欄位)   {len(ca)} 欄 / {len(cb)} 欄", d

    xa, xb, bo = rowset_diff([a], [b])
    if xa == 0 and xb == 0:
        return f"一樣 (只有行序不同)   {len(ca)} 欄 / {na:,} 列", []

    _o1, _o2, pairs = diff_detail([a], [b])
    colc: Counter = Counter()
    real = 0
    ex: list = []
    for ra, rb, d in pairs:
        only_order = bool(d)
        for k in d:
            nm = _cname(ca, k)
            od = _same_multi(ra[k], rb[k])
            if not od:
                only_order = False
                colc[nm] += 1
                if len(ex) < 2 and nm.strip().lower() not in PII:
                    ex.append((nm, ra[k], rb[k]))
            _diffs.append([a.name, b.name, nm, ra[k], rb[k],
                           "欄內順序" if od else "值不同"])
        if not only_order:
            real += 1

    if real == 0:
        return f"一樣 (只有欄內 | 順序不同)   {len(ca)} 欄 / {na:,} 列", []

    det = [f"{nm}: {k:,} 列" for nm, k in colc.most_common(6)]
    det += [f"例 {nm}   {DIR_NEW}={x!r}  {DIR_OLD}={y!r}" for nm, x, y in ex]
    if na != nb:
        return (f"不一樣 (列數 {na:,} / {nb:,}, 差 {abs(na - nb):,})", det)
    return f"不一樣 (值, {real:,} 列)   {len(ca)} 欄 / {na:,} 列", det


def main() -> None:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    ap = argparse.ArgumentParser()
    ap.add_argument("--quick", action="store_true", help="只看檔名與大小 (秒級)")
    ap.add_argument("--only", default=None, help="只做檔名含這段字的")
    args = ap.parse_args()

    root = Path(".").resolve()
    base = root / ROOT_DIRNAME
    if not base.is_dir():
        base = root
    dn, do = base / DIR_NEW, base / DIR_OLD
    for d in (dn, do):
        if not d.is_dir():
            sys.exit(f"找不到 {d}\n  (資料夾名寫在 filecheck.py 最上面的設定區)")

    t0 = time.time()
    add(f"filecheck.py  build {BUILD}")
    add(f"比對 {base}")
    add(f"  新來源 {DIR_NEW}   舊來源 {DIR_OLD}")
    okn, sern, badn = collect(dn, None)
    oko, sero, bado = collect(do, OLD_SERIES)

    add("")
    add("=" * 72)
    add("一、系列盤點")
    add("=" * 72)
    for tag, ser, keep in ((DIR_NEW, sern, None), (DIR_OLD, sero, OLD_SERIES)):
        add("")
        add(f"  {tag}")
        for sn, ps in sorted(ser.items(), key=lambda kv: -len(kv[1])):
            sz = [q.stat().st_size for q in ps]
            ks = sorted({parse_name(q.name)[:2] for q in ps
                         if parse_name(q.name)})
            rng = (f"{ks[0][0]}-{ks[0][1]:02d} ~ {ks[-1][0]}-{ks[-1][1]:02d}"
                   if ks else "?")
            used = (not keep) or keep.lower() in sn
            add(f"    {'比對  ' if used else '不比對'}  {len(ps):>3} 檔  "
                f"{mb(min(sz)):>9}~{mb(max(sz)):<9} {rng}   [{sn}]")
    skipped = [sn for sn in sero if OLD_SERIES.lower() not in sn]
    if skipped:
        add("")
        add(f"  {DIR_OLD} 有 {len(skipped)} 個系列在 {DIR_NEW} 沒有對應的檔, 不配對。")
    for p in badn + bado:
        add(f"    !! 檔名解析不出期間: {p.name}")

    add("")
    add("=" * 72)
    add("二、配對")
    add("=" * 72)
    kn, ko = set(okn), set(oko)
    both, only_n, only_o = sorted(kn & ko), sorted(kn - ko), sorted(ko - kn)
    add(f"  兩邊都有 {len(both)} 個月   只有 {DIR_NEW} {len(only_n)} 個月   "
        f"只有 {DIR_OLD} {len(only_o)} 個月")
    for tag, ks, ok in ((DIR_NEW, only_n, okn), (DIR_OLD, only_o, oko)):
        if ks:
            add("")
            add(f"  只有 {tag} 有:")
            for ym in ks:
                for h in sorted(ok[ym]):
                    for p in ok[ym][h]:
                        add(f"    {ym[0]}-{ym[1]:02d} {HALF[h]:<4} {p.name}")
                        _rec(ym, h, p if tag == DIR_NEW else None,
                             None if tag == DIR_NEW else p, f"只有 {tag} 有")

    add("")
    add("=" * 72)
    add("三、比對結果")
    add("=" * 72)
    v: Counter = Counter()
    same: list = []
    diff: list = []
    for ym in both:
        hn, ho = okn[ym], oko[ym]
        for h in sorted(set(hn) & set(ho)):
            if len(hn[h]) != 1 or len(ho[h]) != 1:
                continue
            a, b = hn[h][0], ho[h][0]
            if args.only and args.only.lower() not in (a.name + b.name).lower():
                continue
            kind, det = compare_file(a, b, args.quick)
            (same if kind.startswith("一樣") else diff).append(
                (ym, h, a, b, kind, det))
            _match[a] = (b.name, kind)
            _match[b] = (a.name, kind)
            v[kind.split(" ")[0]] += 1
            _rec(ym, h, a, b, kind)

    add(f"  一樣的   {len(same)} 組")
    for ym, h, a, b, kind, _d in same:
        add(f"    {ym[0]}-{ym[1]:02d} {HALF[h]}   {kind}")
        add(f"      {DIR_NEW:<7} {a.name}")
        add(f"      {DIR_OLD:<7} {b.name}")
    if not same:
        add("    (無)")
    add("")
    add(f"  不一樣的   {len(diff)} 組")
    for ym, h, a, b, kind, det in diff:
        add(f"    {ym[0]}-{ym[1]:02d} {HALF[h]}   {kind}")
        add(f"      {DIR_NEW:<7} {a.name}")
        add(f"      {DIR_OLD:<7} {b.name}")
        for x in det:
            add(f"        {x}")
    if not diff:
        add("    (無)")

    # 沒有對手的, 只列數量與檔名, 不做內容判斷
    for tag, ok, other in ((DIR_NEW, okn, oko), (DIR_OLD, oko, okn)):
        solo = []
        for ym, hs in ok.items():
            for h, ps in hs.items():
                if ym not in other or h not in other[ym]:
                    for q in ps:
                        solo.append((ym, h, q))
        add("")
        add(f"  只有 {tag} 有   {len(solo)} 個檔 (沒有對手, 無從比對)")
        for ym, h, q in sorted(solo):
            add(f"    {ym[0]}-{ym[1]:02d} {HALF[h]:<4} {q.name}")
            _rec(ym, h, q if tag == DIR_NEW else None,
                 None if tag == DIR_NEW else q, f"只有 {tag} 有")
    skip = [(sn, ps) for sn, ps in sero.items() if OLD_SERIES.lower() not in sn]
    if skip:
        n_ = sum(len(ps) for _, ps in skip)
        add("")
        add(f"  只有 {DIR_OLD} 有 (系列不在配對範圍)   {n_} 個檔")
        for sn, ps in skip:
            add(f"    [{sn}]   {len(ps)} 個檔")

    # ---------- 四、全檔清點 ----------
    # 配對只涵蓋兩邊都有的那幾個月。其餘的檔也要留下 row/column 的紀錄 ——
    # 底稿要的是「我們收到什麼」的完整清單, 不是只有對得上的那幾個。
    inv: list = []
    if not args.quick:
        add("")
        add("=" * 72)
        add("四、全檔清點 (含沒有配對到的)")
        add("=" * 72)
        for tag, ser, d in ((DIR_NEW, sern, dn), (DIR_OLD, sero, do)):
            add("")
            add(f"  {tag}")
            for sn, ps in sorted(ser.items(), key=lambda kv: -len(kv[1])):
                byhdr: dict = {}
                tot = 0
                for q in ps:
                    sha_, hdr_, n_ = scan(q)
                    byhdr.setdefault(hdr_, []).append(q)
                    tot += n_
                    k_ = parse_name(q.name)
                    mt, vd = _match.get(q, ("", "未配對"))
                    dd = sorted(_DAYS.get(q) or {})
                    inv.append([tag, sn, q.name,
                                f"{k_[0]}-{k_[1]:02d}" if k_ else "",
                                HALF.get(k_[2], "") if k_ else "",
                                q.stat().st_size, len(cols_of(hdr_)), n_,
                                dd[0] if dd else "", dd[-1] if dd else "",
                                len(dd), mt, vd, sha_])
                add(f"    [{sn}]   {len(ps)} 檔   合計 {tot:,} 列")
                if len(byhdr) == 1:
                    add(f"      欄位一致: {len(cols_of(next(iter(byhdr))))} 欄")
                    continue
                add(f"      !! 同一個系列內部有 {len(byhdr)} 種欄位版本:")
                base_ = max(byhdr, key=lambda k: len(byhdr[k]))
                bc = cols_of(base_)
                for h_, qs in sorted(byhdr.items(), key=lambda kv: -len(kv[1])):
                    cc = cols_of(h_)
                    note = ""
                    if h_ != base_:
                        ex = [c for c in cc if c not in bc]
                        mi = [c for c in bc if c not in cc]
                        if ex:
                            note += f"  多: {', '.join(ex[:6])}"
                        if mi:
                            note += f"  少: {', '.join(mi[:6])}"
                        if not ex and not mi:
                            note = "  (欄名相同, 順序不同)"
                    add(f"         {len(cc):>3} 欄 × {len(qs):>3} 檔{note}")
                    for q in qs[:2]:
                        add(f"              {q.name}")
                    if len(qs) > 2:
                        add(f"              ... 另外 {len(qs) - 2} 個")

    allf: list = []
    # ---------- 五、逐檔對照 ----------
    # 每一個檔都要有一列, 不是只有配對到的那幾個 —— 底稿是逐檔簽核的。
    add("")
    add("=" * 72)
    for tag, ser in ((DIR_NEW, sern), (DIR_OLD, sero)):
        for sn, ps in ser.items():
            for q in ps:
                k = parse_name(q.name)
                allf.append(((k[0], k[1], k[2]) if k else (9999, 99, 9),
                             tag, sn, q))
    add(f"五、逐檔對照 (全部 {len(allf)} 個檔, 配對的併成一列)")
    add("=" * 72)
    W = 44
    add(f"  {'期間':<9}{'半月':<7}{DIR_NEW:<{W}}{DIR_OLD:<{W}}狀態")
    for key, tag, sn, q in sorted(allf, key=lambda x: (x[0], x[1])):
        y, mo, h = key
        per = f"{y}-{mo:02d}" if y < 9999 else "?"
        mate, verdict = _match.get(q, ("", ""))
        if not verdict:
            verdict = ("未配對 (系列不在範圍)" if tag == DIR_OLD
                       and OLD_SERIES.lower() not in sn else "未配對 (對方無此期間)")
        nm = q.name[:W - 2]
        cn = nm if tag == DIR_NEW else (mate[:W - 2] if mate else "—")
        on = nm if tag == DIR_OLD else (mate[:W - 2] if mate else "—")
        if mate and tag == DIR_OLD:
            continue          # 已經由對手那一列印過了
        add(f"  {per:<9}{HALF[h]:<5}{cn:<{W}}{on:<{W}}{verdict}")

    # ---------- 六、期間覆蓋 ----------
    # 這一節對每一個檔都成立, 不需要對手 —— 沒有對照組的檔也能檢查:
    # 檔名宣稱的期間, 內容有沒有真的對上。多一天少一天都是要問客戶的。
    cov: list = []
    if not args.quick:
        add("")
        add("=" * 72)
        add("六、期間覆蓋 (檔名宣稱的期間 vs 內容實際的日期)")
        add("=" * 72)
        add(f"  {'期間':<9}{'半月':<6}{'來源':<8}{'內容日期範圍':<26}{'天數':>5}  判定")
        nbad = 0
        for key, tag, sn, q in sorted(allf, key=lambda x: (x[0], x[1])):
            y, mo, h = key
            days = _DAYS.get(q) or {}
            per = f"{y}-{mo:02d}" if y < 9999 else "?"
            if not days:
                add(f"  {per:<9}{HALF[h]:<4}{tag:<8}(找不到日期欄)")
                cov.append([tag, q.name, per, HALF[h], "", "", 0, "無日期欄"])
                continue
            ds = sorted(days)
            lo, hi = ds[0], ds[-1]
            # 檔名說上半月就該落在 1~15, 下半月 16 以後, 整月不限
            out = []
            for d in ds:
                if len(d) < 10 or not d.startswith(f"{y}-{mo:02d}"):
                    out.append(d)
                elif h == 1 and int(d[8:10]) > 15:
                    out.append(d)
                elif h == 2 and int(d[8:10]) <= 15:
                    out.append(d)
            # 該有幾天: 上半月 15, 下半月 該月天數-15, 整月 該月天數
            mdays = calendar.monthrange(y, mo)[1] if y < 9999 else 30
            want = 15 if h == 1 else (mdays - 15 if h == 2 else mdays)
            inr = [d for d in ds if d not in out]
            miss = want - len(inr)
            notes = []
            if out:
                notes.append(f"{len(out)} 天在期間外 ({out[0]}~{out[-1]})")
            if miss > 0:
                notes.append(f"期間內缺 {miss} 天")
            if notes:
                nbad += 1
                verdict = "!! " + ", ".join(notes)
            else:
                verdict = "符合"
            add(f"  {per:<9}{HALF[h]:<4}{tag:<8}{lo} ~ {hi:<14}{len(ds):>5}  {verdict}")
            cov.append([tag, q.name, per, HALF[h], lo, hi, len(ds),
                        verdict.replace("!! ", "")])
        add("")
        add(f"  {len(cov) - nbad} 個檔內容與檔名相符, {nbad} 個不符。")

    # ---------- 七、合併比對 ----------
    # 兩邊各自把同一個月的所有檔併起來 (含沒有配對的系列) 再比。
    #
    # 為什麼不是整批全合: client 1,600 萬列 + ita 1,166 萬列, 鍵集合會吃掉
    # 1GB 以上。而第六節已經證明每個檔的內容就是檔名說的那個月, 所以日期
    # 不重疊的月份不可能有共同記錄 —— 只合有交集的月份, 結果跟全合一樣。
    #
    # 用 票號+航班日期+航班號 當鍵, 不用整行雜湊 —— 兩邊欄位版本不同
    # (54 / 56 / 19 欄), 整行比必然全部對不上, 那不是資料差異。
    if not args.quick:
        add("")
        add("=" * 72)
        add("七、合併比對 (同月份的檔全部併起來, 含沒有配對的系列)")
        add("=" * 72)
        mn: dict = {}
        mo_: dict = {}
        for tgt, ser in ((mn, sern), (mo_, sero)):
            for ps in ser.values():
                for q in ps:
                    k = parse_name(q.name)
                    if k:
                        tgt.setdefault(k[:2], []).append(q)
        shared = sorted(set(mn) & set(mo_))
        add(f"  {DIR_NEW} 涵蓋 {len(mn)} 個月, {DIR_OLD} 涵蓋 {len(mo_)} 個月, "
            f"兩邊都有 {len(shared)} 個月")
        if not shared:
            add("  沒有共同月份 —— 兩批資料的期間完全錯開。")
        gb = gn = go = 0
        for ym in shared:
            fa, fb = mn[ym], mo_[ym]
            ka, kb = set(), set()
            okk = True
            for ps, dst in ((fa, ka), (fb, kb)):
                for q in ps:
                    t = _keyset(q)
                    if t is None:
                        okk = False
                        add(f"  {ym[0]}-{ym[1]:02d}  !! {q.name} 找不到"
                            " 票號/日期/航班號 欄, 這個月無法比")
                        break
                    dst |= t
            if not okk:
                continue
            b_ = len(ka & kb)
            gb += b_
            gn += len(ka - kb)
            go += len(kb - ka)
            add("")
            add(f"  {ym[0]}-{ym[1]:02d}   {DIR_NEW} {len(fa)} 檔 {len(ka):,} 筆"
                f" / {DIR_OLD} {len(fb)} 檔 {len(kb):,} 筆")
            add(f"    共同 {b_:,}   只在 {DIR_NEW} {len(ka - kb):,}"
                f"   只在 {DIR_OLD} {len(kb - ka):,}")
            if b_ and not (kb - ka):
                add(f"    -> {DIR_OLD} 這個月的記錄, {DIR_NEW} 全部都有")
            elif b_ and not (ka - kb):
                add(f"    -> {DIR_NEW} 這個月的記錄, {DIR_OLD} 全部都有")
            elif not b_:
                add("    -> 完全沒有共同記錄")
        solo_n = sorted(set(mn) - set(mo_))
        solo_o = sorted(set(mo_) - set(mn))
        add("")
        add("  合計 (只算有共同月份的部分, 其餘月份日期不重疊, 不可能相同)")
        add(f"    共同的記錄      {gb:,}")
        add(f"    只在 {DIR_NEW}   {gn:,}   另有 {len(solo_n)} 個月 "
            f"{DIR_OLD} 完全沒有")
        add(f"    只在 {DIR_OLD}   {go:,}   另有 {len(solo_o)} 個月 "
            f"{DIR_NEW} 完全沒有")

    # ---------- 八、交叉驗證 ----------
    # 上面的配對建立在「檔名怎麼解析成期間」這條規則上。規則錯了整份報告
    # 就白做, 所以用一個完全不看檔名的方法驗一次: 位元組數一樣的檔幾乎
    # 一定是同一個檔。大小相同卻沒配到 -> 期間規則有問題。
    add("")
    add("=" * 72)
    add("八、交叉驗證 (不看檔名, 只比位元組數)")
    add("=" * 72)
    szn: dict = {}
    for ps in sern.values():
        for q in ps:
            szn.setdefault(q.stat().st_size, []).append(q)
    same_sz = []
    for ps in sero.values():
        for q in ps:
            for w in szn.get(q.stat().st_size, []):
                same_sz.append((q.stat().st_size, w, q))
    add(f"  位元組數完全相同的組合: {len(same_sz)}")
    bad_map = 0
    for n_, (sz, w, q) in enumerate(sorted(same_sz)):
        ok_ = _match.get(w, ("", ""))[0] == q.name
        if not ok_:
            bad_map += 1
        if n_ < 40:      # 大小大量碰撞時不要洗版
            # 這批檔名共用很長的前綴, 截頭會全都長一樣 —— 留尾巴才分得出
            add(f"    {sz:>15,}   {tail(w.name):<34}{tail(q.name)}"
                + ("" if ok_ else "   !! 沒有被配到"))
    if len(same_sz) > 40:
        add(f"    ... 另外 {len(same_sz) - 40} 組")
    if bad_map:
        add(f"  !! 有 {bad_map} 組大小相同卻沒配到 —— 期間解析規則要檢查")
    else:
        add("  全部大小相同的組合都有被配到, 期間解析規則與檔案內容一致。")

    add("")
    add("=" * 72)
    add("總結 (檔數)")
    add("=" * 72)
    tn = sum(len(x) for x in sern.values())
    to = sum(len(x) for x in sero.values())
    npair = sum(v.values())
    add(f"    {DIR_NEW} 共 {tn} 檔   {DIR_OLD} 共 {to} 檔")
    add(f"    有對手的 {npair} 組 ({npair * 2} 檔):")
    for k, n in sorted(v.items()):
        add(f"      {k:<8}{n:>4} 組 ({n * 2} 檔)")
    add(f"    沒有對手: {DIR_NEW} {tn - npair} 檔   {DIR_OLD} {to - npair} 檔")
    add("")
    add(f"完成, 共 {time.time() - t0:.0f}s")

    out = root / REPORT
    out.write_text("\n".join(_lines), encoding="utf-8")
    sm = root / SUMMARY
    with open(sm, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["期間", "半月", f"{DIR_NEW}檔名", f"{DIR_OLD}檔名",
                    f"{DIR_NEW}大小", f"{DIR_OLD}大小", "SHA256",
                    f"{DIR_NEW}欄數", f"{DIR_OLD}欄數", "欄位相同",
                    f"{DIR_NEW}列數", f"{DIR_OLD}列數", "列數相同",
                    f"只在{DIR_NEW}的列", f"只在{DIR_OLD}的列", "判定"])
        w.writerows(_rows)
    print(f"\n報告: {out}")
    print(f"底稿表: {sm}")
    print("(這兩份只有檔名與統計值, 沒有任何一列資料內容)")
    if inv:
        ip = root / INVENTORY
        with open(ip, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(["來源", "系列", "檔名", "期間", "半月", "大小",
                        "欄數", "列數", "最早日期", "最晚日期", "涵蓋天數",
                        "配對到的檔", "狀態", "SHA256"])
            w.writerows(inv)
        print(f"全檔清單: {ip}   ({len(inv)} 個檔)")
    if _diffs:
        dp = root / DIFFS
        with open(dp, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow([f"{DIR_NEW}檔名", f"{DIR_OLD}檔名", "欄名",
                        f"{DIR_NEW}的值", f"{DIR_OLD}的值", "類型"])
            w.writerows(_diffs)
        print(f"差異明細: {dp}   ({len(_diffs):,} 筆)")
        print("⚠️ 這一份含原始欄位值, 可能有個資, 走內部管道")


if __name__ == "__main__":
    main()
