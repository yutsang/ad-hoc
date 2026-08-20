#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
產出 Power BI 用的星型結構 (長格式), 放在 pbi/ 底下。

模板那個 131 欄的寬表是給人看的, 不適合 Power BI: 126 個月份欄當不了時間軸,
也做不了時間智慧。這裡改成一張事實表 + 三張維度表。

門店的城市/公司是從明細 parquet 取的 (聚合表裡沒有), 這樣就能按城市、
公司切片, 而不是只能看門店編碼。

用法:
    python export_powerbi.py
"""
import sys
from pathlib import Path

import duckdb

ROOT = Path(__file__).resolve().parent
PQ = ROOT / "parquet"
AGG = ROOT / "aggregate.parquet"
OUT = ROOT / "pbi"


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    if not AGG.exists():
        sys.exit(f"找不到 {AGG.name}, 先跑 build_aggregate.py")
    # 聚合規則改過就得重跑, 不然這裡匯出的是舊名稱, 而且不會有任何徵兆
    builder = ROOT / "build_aggregate.py"
    if builder.exists() and AGG.stat().st_mtime < builder.stat().st_mtime:
        sys.exit(f"{AGG.name} 比 {builder.name} 舊, 是用舊規則產生的。\n"
                 f"請先重跑:  python build_aggregate.py --channel-map channel_map.csv")
    OUT.mkdir(exist_ok=True)

    con = duckdb.connect()
    agg = f"read_parquet('{AGG.as_posix()}')"
    det = f"read_parquet('{PQ.as_posix()}/*.parquet')"

    # 商品编码 必須唯一決定 商品名称 與 产品分类, 否則不能搬進維度表
    bad = con.execute(f"""
        SELECT count(*) FROM (
            SELECT "商品编码" FROM {agg} GROUP BY 1
            HAVING count(DISTINCT COALESCE("商品名称",'')) > 1
                OR count(DISTINCT COALESCE("产品分类",'')) > 1)
    """).fetchone()[0]
    if bad:
        sys.exit(f"有 {bad} 個商品编码 對到多個名稱或分類, 不能建維度表")

    def dump(name, sql):
        p = OUT / f"{name}.parquet"
        con.execute(f"COPY ({sql}) TO '{p.as_posix()}' (FORMAT PARQUET, COMPRESSION ZSTD)")
        n = con.execute(f"SELECT count(*) FROM read_parquet('{p.as_posix()}')").fetchone()[0]
        print(f"  {name:<12} {n:>10,} 列  {p.stat().st_size/1e6:>7,.1f} MB")

    print("寫出 pbi/ ...")

    # 事實表: 期間轉成真正的日期, Power BI 的時間智慧才能用
    dump("销售事实", f"""
        SELECT "门店编码", "销售渠道", "商品编码",
               CAST("period" || '-01' AS DATE) AS "日期",
               "商品数量", "应收金额", "实收金额", "明细列数"
        FROM {agg}
    """)

    dump("商品维度", f"""
        SELECT "商品编码", any_value("商品名称") AS "商品名称",
               any_value("产品分类") AS "产品分类"
        FROM {agg} GROUP BY 1
    """)

    # 門店的城市/公司只有明細有; 門店可能改過名, 取最後一次出現的
    dump("门店维度", f"""
        SELECT "门店编码",
               arg_max("门店名称", "period") FILTER (WHERE "门店名称" IS NOT NULL) AS "门店名称",
               arg_max("城市",     "period") FILTER (WHERE "城市"     IS NOT NULL) AS "城市",
               arg_max("公司",     "period") FILTER (WHERE "公司"     IS NOT NULL) AS "公司"
        FROM {det} WHERE "门店编码" IS NOT NULL GROUP BY 1
    """)

    # 季度用內建 quarter(), 別自己算 —— (月+2)/3 再 CAST 會四捨五入, 3 月會變第 2 季
    dump("日期维度", f"""
        SELECT DISTINCT
               CAST("period" || '-01' AS DATE)                       AS "日期",
               year(CAST("period" || '-01' AS DATE))                 AS "年",
               month(CAST("period" || '-01' AS DATE))                AS "月",
               "period"                                              AS "年月",
               '第' || quarter(CAST("period" || '-01' AS DATE)) || '季' AS "季度"
        FROM {agg} ORDER BY 1
    """)

    # 對帳: 事實表總數必須等於 aggregate
    a = con.execute(f"""SELECT sum("商品数量"), sum("应收金额"), sum("实收金额") FROM {agg}""").fetchone()
    b = con.execute(f"""SELECT sum("商品数量"), sum("应收金额"), sum("实收金额")
                        FROM read_parquet('{(OUT/"销售事实.parquet").as_posix()}')""").fetchone()
    print("\n對帳 (事實表 vs aggregate.parquet):")
    ok = True
    for m, x, y in zip(["商品数量", "应收金额", "实收金额"], a, b):
        d = (x or 0) - (y or 0)
        if abs(d) > 0.005:
            ok = False
        print(f"  {m}: {x:,.2f} vs {y:,.2f}   差 {d:,.2f}")
    print(f"  -> {'對帳通過' if ok else '!! 對帳失敗'}")
    if not ok:
        sys.exit(1)


if __name__ == "__main__":
    main()
