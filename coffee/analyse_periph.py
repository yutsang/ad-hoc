#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
勘查寬表裡「周边产品」這個維度: 分類到底有幾個、攤平後幾列、資料夠不夠用。

一趟回答三件事:
  1. 寬表實際有哪些 产品分类 —— 跟項目組給的名單逐字比對。名單外的、名單有
     但資料沒有的、長得像但字串不一樣的 (全半形/空白), 全部列出來。
     順便把分類開頭的數字碼排出來, 名單跳號的地方 (11-14 之後直接 16) 一看就知道。
  2. 照 build_sku_monthly.py 的規則攤成月度長表之後, 每個分類各幾列。
     产品分类 在 group key 裡, 所以任選一組分類的列數就是相加 —— 名單怎麼改
     都不必重跑。順便印上次交的 01/02 幾列, 拿來對照這支算得對不對。
  3. sanity check: 月份覆蓋、空名稱、同名不同編碼、渠道合併、負值、逐類對帳。

輸入是 inspect_files.py 產的 宽表.parquet, 不碰 105MB 的 xlsx, 整趟幾十秒。
輸出全是聚合統計、不含明細列 (除非加 --names), 報告可以直接貼出來討論。

用法:
    python analyse_periph.py
    python analyse_periph.py --root D:\\某資料夾
    python analyse_periph.py --names          # 額外印每類前幾個商品名称 (含明細值)
"""
from __future__ import annotations

import argparse
import re
import sys
import time
import unicodedata
from pathlib import Path

try:
    import duckdb
except ImportError:
    sys.exit("需要 duckdb:  pip install duckdb")

# ================================================================= 設定
PQ = "宽表.parquet"                  # inspect_files.py 的產物
REPORT = "周边勘查.txt"

MEASURES = ["商品数量", "应收金额", "实收金额"]
GROUP_KEYS = ["月份", "门店编码", "产品分类", "商品名称"]   # 模板粒度, 同 build_sku_monthly.py

# 項目組給的周邊名單 (照原樣, 不排序)
WANT = [
    "11餐盒",
    "12周边产品-杯子碗碟",
    "13周边产品-服装配饰",
    "14周边产品-包袋挂件",
    "17周边产品-生活用品",
    "18周边产品-其他货品",
    "19周边产品-特别限定门店",
    "16周边产品-咖啡器具",
]
DONE = ["01咖啡类饮品", "02非咖啡类饮品"]   # 上次已交付的, 當對照基準
PERIPH_HINT = ("周边", "周邊", "餐盒")       # 名單外但長這樣的, 視為候選周邊

EXCEL_MAX_ROWS = 1_048_576
TPL_HEAD_ROWS = 5                    # 模板前 5 列是抬頭
TOP_NAMES = 5                        # --names 時每類印幾個
# ================================================================= 設定結束

L: list[str] = []


def add(s: str = "") -> None:
    L.append(s)
    print(s, flush=True)


def q(name: str) -> str:
    return '"' + name.replace('"', '""') + '"'


def lit(v: str) -> str:
    return "'" + v.replace("'", "''") + "'"


def width(t: str) -> int:
    return sum(2 if ord(c) > 0x2E80 else 1 for c in t)


def pad(t: str, w: int) -> str:
    """中文字佔兩格, ljust 會對不齊, 自己算顯示寬度。"""
    return t + " " * max(0, w - width(t))


def rj(t: str, w: int) -> str:
    """靠右版的 pad —— 中文表頭要跟數字欄對齊只能自己算。"""
    return " " * max(0, w - width(t)) + t


def norm(t: str) -> str:
    """比對用的正規化: 全形轉半形、去空白、大小寫一致 —— 專抓「看起來一樣」的字串。"""
    t = unicodedata.normalize("NFKC", t or "")
    return "".join(c for c in t if not c.isspace()).lower()


def code_of(cat: str) -> str:
    """分類開頭的數字碼, 例如 '12周边产品-杯子碗碟' -> '12'。沒有就回空字串。"""
    m = re.match(r"\s*(\d+)", cat or "")
    return m.group(1) if m else ""


def parse_cols(cols: list[str]) -> tuple[dict, list[str]]:
    """寬表欄名長這樣: '"商品数量"按“营业日期”加总|2023-01' -> {度量: {月份: 欄名}}"""
    out: dict[str, dict[str, str]] = {m: {} for m in MEASURES}
    for c in cols:
        if "|" not in c:
            continue
        title, period = c.rsplit("|", 1)
        for m in MEASURES:
            if m in title:
                out[m][period] = c
                break
    periods = sorted(out[MEASURES[0]])
    for m in MEASURES:
        if sorted(out[m]) != periods:
            sys.exit(f"{m} 的月份欄跟 {MEASURES[0]} 對不起來, "
                     f"{len(out[m])} vs {len(periods)} 個")
    if not periods:
        sys.exit(f"{PQ} 裡找不到 '度量|月份' 形式的欄位, 這不是 inspect_files.py 的產物")
    return out, periods


# --------------------------------------------------- 1. 分類清單 vs 名單
def inventory(con, src: str, cmap: dict, periods: list[str]) -> list[dict]:
    """寬表全域的分類分佈。三個度量是把該分類所有月份的欄橫向加起來。"""
    tot = ", ".join(
        "sum(" + " + ".join(f"COALESCE({q(cmap[m][p])}, 0)" for p in periods)
        + f") AS {q(m)}" for m in MEASURES)
    rows = con.sql(f'''
        SELECT COALESCE("产品分类", '(空白)') AS cat, count(*) AS n,
               count(DISTINCT "门店编码"), count(DISTINCT "商品编码"),
               count(DISTINCT "商品名称"), {tot}
        FROM {src} GROUP BY 1 ORDER BY 2 DESC''').fetchall()
    return [{"cat": r[0], "n": r[1], "store": r[2], "code": r[3], "name": r[4],
             "sums": list(r[5:])} for r in rows]


def report_inventory(inv: list[dict]) -> tuple[list[str], list[str]]:
    """印出分類清單並跟名單比對, 回傳 (資料裡真的有的名單分類, 名單外的候選周邊)。"""
    have = {r["cat"]: r for r in inv}
    nmap = {norm(r["cat"]): r["cat"] for r in inv}
    done_set = set(DONE)

    # 先配對再印表 —— 不然全半形不同的那類會在表上被誤標成「名單外」。
    hit, miss, near = [], [], {}
    for c in WANT:
        if c in have:
            hit.append(c)
        elif norm(c) in nmap:
            hit.append(nmap[norm(c)])
            near[nmap[norm(c)]] = c
        else:
            miss.append(c)
    hit_set = set(hit)
    extra = [r["cat"] for r in inv
             if r["cat"] not in hit_set | done_set
             and any(h in r["cat"] for h in PERIPH_HINT)]

    w = max([width(r["cat"]) for r in inv] + [12])
    add(f"=== 寬表裡的 产品分类 共 {len(inv)} 個 ===")
    add(f"  {pad('分类', w)}  {rj('寬表列数', 12)}  {rj('门店', 6)}  "
        f"{rj('商品编码', 10)}  {rj('商品名称', 10)}  {rj(MEASURES[0], 16)}  "
        f"{rj(MEASURES[1], 18)}  {rj(MEASURES[2], 18)}")
    for r in inv:
        if r["cat"] in near:
            tag = f"<- 名單 (名單寫的是 {near[r['cat']]!r}, 字串不同)"
        elif r["cat"] in hit_set:
            tag = "<- 名單"
        elif r["cat"] in done_set:
            tag = "   (上次已交)"
        elif any(h in r["cat"] for h in PERIPH_HINT):
            tag = "!! 名單外, 但看起來是周邊"
        else:
            tag = ""
        add(f"  {pad(r['cat'], w)}  {r['n']:>12,}  {r['store']:>6,}  {r['code']:>10,}  "
            f"{r['name']:>10,}  {r['sums'][0]:>16,.0f}  {r['sums'][1]:>18,.2f}  "
            f"{r['sums'][2]:>18,.2f}  {tag}")

    # ---- 名單逐條對照 ----
    add("")
    add(f"=== 名單 {len(WANT)} 類逐條對照 ===")
    for c in WANT:
        if c in have:
            add(f"  [有]   {c}   {have[c]['n']:,} 列")
        elif norm(c) in nmap:
            real = nmap[norm(c)]
            add(f"  [近似] {c}  -> 資料裡是 {real!r} (全半形或空白不同), "
                f"{have[real]['n']:,} 列")
        else:
            add(f"  !! [缺]  {c}   寬表裡完全沒有這個字串")

    add("")
    if extra:
        add(f"=== !! 名單外、但名稱像周邊的分類 {len(extra)} 個 ===")
        for c in extra:
            add(f"  {c}   {have[c]['n']:,} 列   "
                f"{MEASURES[0]} {have[c]['sums'][0]:,.0f}")
        add("  -> 要不要一起出, 需要項目組確認")
    else:
        add("=== 名單外沒有名稱像周邊的分類 ===")

    # ---- 數字碼覆蓋: 名單跳號的地方最容易漏 ----
    codes = {}
    for r in inv:
        c = code_of(r["cat"])
        if c:
            codes.setdefault(c, []).append(r["cat"])
    add("")
    add("=== 分類數字碼 (資料裡實際有的) ===")
    for c in sorted(codes):
        mark = "" if any(x in set(hit) | done_set for x in codes[c]) else "   <- 不在名單/上次範圍"
        add(f"  {c}: {', '.join(codes[c])}{mark}")
    want_codes = {code_of(c) for c in WANT} - {""}
    if want_codes:
        lo, hi = min(want_codes), max(want_codes)
        gap = [c for c in sorted(codes)
               if lo <= c <= hi and c not in want_codes
               and c not in {code_of(x) for x in DONE}]
        if gap:
            add(f"  !! 名單碼段 {lo}~{hi} 之間, 資料裡還有名單沒列到的碼: "
                f"{', '.join(gap)}")
    if miss:
        add("")
        add(f"  !! 名單有 {len(miss)} 類在寬表裡找不到: {miss}")
    return hit, extra


# --------------------------------------------------- 2. 攤平後的列數
def flatten(con, src: str, cmap: dict, periods: list[str]) -> str:
    """照 build_sku_monthly.py 的規則攤平 + 聚合到模板粒度, 但不篩分類。

    三個度量全空的月份不產生列 —— 這條規則決定了最後的列數, 跟正式那支一致。
    """
    parts = []
    for p in periods:
        vals = ", ".join(f"{q(cmap[m][p])} AS {q(m)}" for m in MEASURES)
        nn = " OR ".join(f"{q(cmap[m][p])} IS NOT NULL" for m in MEASURES)
        parts.append(f'SELECT "门店编码", "产品分类", "商品名称", {lit(p)} AS "月份",'
                     f' {vals} FROM {src} WHERE {nn}')
    long_sql = "\nUNION ALL\n".join(parts)
    keys = ", ".join(q(k) for k in GROUP_KEYS)
    sums = ", ".join(f'sum(COALESCE({q(m)}, 0)) AS {q(m)}' for m in MEASURES)
    con.execute(f"""
        CREATE TEMP TABLE agg AS
        SELECT {keys}, {sums}
        FROM ({long_sql})
        GROUP BY {", ".join(str(i + 1) for i in range(len(GROUP_KEYS)))}
    """)
    return "agg"


def report_rows(con, agg: str, targets: list[str]) -> int:
    """每個分類攤平後幾列 —— 分類在 group key 裡, 所以子集列數 = 相加。"""
    rows = con.sql(f'''
        SELECT COALESCE("产品分类", '(空白)') AS cat, count(*) AS n,
               count(DISTINCT "门店编码"), count(DISTINCT "商品名称"),
               min("月份"), max("月份"), count(DISTINCT "月份"),
               count(*) FILTER (WHERE "商品名称" IS NULL OR "商品名称" = '')
        FROM {agg} GROUP BY 1 ORDER BY 2 DESC''').fetchall()
    w = max([width(r[0]) for r in rows] + [12])
    tset, dset = set(targets), set(DONE)
    add(f"=== 攤成月度長表 (月份×门店×分类×商品名称) 之後的列數 ===")
    add(f"  {pad('分类', w)}  {rj('攤平列数', 12)}  {rj('门店', 6)}  "
        f"{rj('商品名称', 10)}  {rj('首月', 8)}  {rj('末月', 8)}  "
        f"{rj('有资料月数', 10)}  {rj('空名称列', 10)}")
    n_target = 0
    for cat, n, st, nm, lo, hi, nmo, blank in rows:
        tag = " <- 目標" if cat in tset else ("   (上次已交)" if cat in dset else "")
        if cat in tset:
            n_target += n
        add(f"  {pad(cat, w)}  {n:>12,}  {st:>6,}  {nm:>10,}  {lo:>8}  {hi:>8}  "
            f"{nmo:>10}  {blank:>10,}{tag}")

    n_done = sum(n for cat, n, *_ in rows if cat in dset)
    cap = EXCEL_MAX_ROWS - TPL_HEAD_ROWS
    add("")
    add(f"=== 列數 vs Excel 上限 ({cap:,} 列 = {EXCEL_MAX_ROWS:,} 扣 {TPL_HEAD_ROWS} 列抬頭) ===")
    add(f"  目標 {len(targets)} 類           {n_target:>12,} 列   "
        f"{'放得下' if n_target <= cap else f'放不下, 超出 {n_target - cap:,} 列'}")
    add(f"  上次的 01/02          {n_done:>12,} 列   (拿來對照上次交出去的數字)")
    add(f"  兩者合併在同一個檔     {n_target + n_done:>12,} 列   "
        f"{'放得下' if n_target + n_done <= cap else f'放不下, 超出 {n_target + n_done - cap:,} 列'}")

    if targets:
        add("")
        add("  目標分類逐年列數:")
        for y, n in con.sql(f'''
                SELECT left("月份", 4), count(*) FROM {agg}
                WHERE "产品分类" IN ({", ".join(lit(c) for c in targets)})
                GROUP BY 1 ORDER BY 1''').fetchall():
            add(f"    {y} 年: {n:,} 列")
    return n_target


# --------------------------------------------------- 3. sanity check
def sanity(con, src: str, agg: str, targets: list[str], cmap: dict,
           periods: list[str]) -> None:
    if not targets:
        add("=== sanity check: 目標分類是空的, 跳過 ===")
        return
    in_t = ", ".join(lit(c) for c in targets)

    # -- 攤平結果的資料品質 --
    n, blank, neg_qty, neg_amt, zero, over = con.sql(f'''
        SELECT count(*),
               count(*) FILTER (WHERE "商品名称" IS NULL OR "商品名称" = ''),
               count(*) FILTER (WHERE "商品数量" < 0),
               count(*) FILTER (WHERE "应收金额" < 0 OR "实收金额" < 0),
               count(*) FILTER (WHERE "商品数量" = 0 AND "应收金额" = 0
                                  AND "实收金额" = 0),
               count(*) FILTER (WHERE "实收金额" > "应收金额")
        FROM {agg} WHERE "产品分类" IN ({in_t})''').fetchone()
    add("=== 攤平結果的資料品質 (目標分類) ===")
    add(f"  總列數 {n:,}")
    add(f"  商品名称 是空的        {blank:,} 列 ({blank / n * 100 if n else 0:.2f}%)"
        + ("   <- 這些列在成果檔會是空白名稱, 跟上次同一個問題" if blank else ""))
    add(f"  商品数量 < 0           {neg_qty:,} 列   (退貨/沖銷, 有就正常)")
    add(f"  应收或实收 < 0         {neg_amt:,} 列")
    add(f"  三個度量全 0           {zero:,} 列   (有值但都是 0, 不是空)")
    add(f"  实收 > 应收            {over:,} 列   (多半是加價購/湊單, 要問業務)")

    # -- 模板沒有的兩個欄位: 掉了會被加總 --
    add("")
    add("=== 模板粒度會合併掉什麼 (目標分類) ===")
    wide_n, out_n = con.sql(f'''
        SELECT count(*), count(DISTINCT ("门店编码", "产品分类", "商品名称"))
        FROM {src} WHERE "产品分类" IN ({in_t})''').fetchone()
    add(f"  寬表 {wide_n:,} 列 -> 去掉 销售渠道/商品编码 之後剩 {out_n:,} 組 key"
        f"   (平均 {wide_n / out_n if out_n else 0:.2f} 列併 1 組)")
    for ch, c in con.sql(f'''
            SELECT COALESCE("销售渠道", '(空白)'), count(*) FROM {src}
            WHERE "产品分类" IN ({in_t}) GROUP BY 1 ORDER BY 2 DESC''').fetchall():
        add(f"    销售渠道 {ch}: {c:,} 列")
    dup_name, dup_rows = con.sql(f'''
        SELECT count(*), COALESCE(sum(n), 0) FROM (
            SELECT "产品分类", "商品名称", count(DISTINCT "商品编码") AS k,
                   count(*) AS n
            FROM {src} WHERE "产品分类" IN ({in_t})
            GROUP BY 1, 2 HAVING k > 1)''').fetchone()
    add(f"  同一分類裡同名不同 商品编码: {dup_name:,} 個名稱, 涉及 {dup_rows:,} 寬表列"
        + ("   <- 這些會被加總成一列" if dup_name else ""))

    # -- 月份覆蓋 --
    add("")
    add(f"=== 月份覆蓋 (寬表共 {len(periods)} 個月: {periods[0]} ~ {periods[-1]}) ===")
    got = {r[0]: r[1] for r in con.sql(
        f'SELECT "月份", count(*) FROM {agg} WHERE "产品分类" IN ({in_t}) '
        f"GROUP BY 1").fetchall()}
    holes = [p for p in periods if p not in got]
    if holes:
        add(f"  !! 目標分類完全沒有資料的月份 {len(holes)} 個: "
            f"{', '.join(holes[:24])}" + (" ..." if len(holes) > 24 else ""))
    else:
        add(f"  {len(periods)} 個月每個月都有資料")
    thin = sorted((c, p) for p, c in got.items())[:5]
    if thin:
        add("  列數最少的幾個月: " + ", ".join(f"{p} {c:,} 列" for c, p in thin))

    # -- 逐類對帳: 橫向 (寬表) vs 縱向 (攤平後) --
    add("")
    add("=== 逐類對帳: 寬表橫向加總 vs 攤平後縱向加總 ===")
    tot = ", ".join(
        "sum(" + " + ".join(f"COALESCE({q(cmap[m][p])}, 0)" for p in periods)
        + ")" for m in MEASURES)
    want = {r[0]: r[1:] for r in con.sql(
        f'SELECT "产品分类", {tot} FROM {src} '
        f"WHERE \"产品分类\" IN ({in_t}) GROUP BY 1").fetchall()}
    have = {r[0]: r[1:] for r in con.sql(
        f'SELECT "产品分类", {", ".join(f"sum({q(m)})" for m in MEASURES)} '
        f"FROM {agg} WHERE \"产品分类\" IN ({in_t}) GROUP BY 1").fetchall()}
    bad = 0
    for c in targets:
        wv, hv = want.get(c, (0.0,) * 3), have.get(c, (0.0,) * 3)
        for i, m in enumerate(MEASURES):
            if abs((wv[i] or 0) - (hv[i] or 0)) > 0.005:
                bad += 1
                add(f"  !! {c} {m}: 寬表 {wv[i]:,.2f} vs 攤平後 {hv[i]:,.2f}")
    add(f"  {len(targets)} 類 × {len(MEASURES)} 個度量 = {len(targets) * len(MEASURES)} 項, "
        f"不符 {bad} 項" + ("   -> 攤平邏輯沒問題" if not bad else "   -> 有問題, 先別出檔"))
    for i, m in enumerate(MEASURES):
        add(f"  目標分類 {m} 總計 {sum(v[i] or 0 for v in want.values()):,.2f}")


def show_names(con, agg: str, targets: list[str]) -> None:
    add("")
    add(f"=== 每類 商品数量 前 {TOP_NAMES} 名的 商品名称 (含明細值, 貼出去前先看一眼) ===")
    for c in targets:
        rows = con.sql(f'''
            SELECT COALESCE("商品名称", '(空白)'), sum("商品数量")
            FROM {agg} WHERE "产品分类" = {lit(c)}
            GROUP BY 1 ORDER BY 2 DESC NULLS LAST LIMIT {TOP_NAMES}''').fetchall()
        add(f"  {c}:")
        for nm, qty in rows:
            add(f"    {nm}   {qty or 0:,.0f}")


# ------------------------------------------------------------------ main
def main() -> None:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

    ap = argparse.ArgumentParser()
    ap.add_argument("--root", default=".", help="資料夾 (預設: 目前目錄)")
    ap.add_argument("--names", action="store_true",
                    help=f"額外印每類前 {TOP_NAMES} 個商品名称 (含明細值)")
    args = ap.parse_args()

    root = Path(args.root).resolve()
    src_pq = root / PQ
    if not src_pq.exists():
        sys.exit(f"找不到 {src_pq}  (先跑 inspect_files.py)")
    t0 = time.time()

    con = duckdb.connect()
    src = f"read_parquet('{src_pq.as_posix()}')"
    cols = [r[0] for r in con.sql(f"DESCRIBE SELECT * FROM {src}").fetchall()]
    cmap, periods = parse_cols(cols)
    n_wide = con.sql(f"SELECT count(*) FROM {src}").fetchone()[0]
    add(f"來源: {src_pq}")
    add(f"寬表 {n_wide:,} 列 × {len(cols)} 欄 = {len(MEASURES)} 個度量 × "
        f"{len(periods)} 個月 ({periods[0]} ~ {periods[-1]})")
    add("")

    inv = inventory(con, src, cmap, periods)
    hit, extra = report_inventory(inv)

    add("")
    print("攤平中 (幾十秒)...", file=sys.stderr, flush=True)
    agg = flatten(con, src, cmap, periods)

    # 目標 = 名單裡資料真的有的 + 名單外看起來是周邊的。後者只是候選, 但先一起算,
    # 免得項目組確認完又要重跑一次。
    targets = hit + extra
    add(f"(目標分類 {len(targets)} 個 = 名單命中 {len(hit)} + 名單外候選 {len(extra)})")
    add("")
    report_rows(con, agg, targets)
    add("")
    sanity(con, src, agg, targets, cmap, periods)
    if args.names:
        show_names(con, agg, targets)

    out = root / REPORT
    out.write_text("\n".join(L), encoding="utf-8")
    print(f"\n報告: {out}   ({time.time() - t0:.0f}s)", file=sys.stderr)


if __name__ == "__main__":
    main()
