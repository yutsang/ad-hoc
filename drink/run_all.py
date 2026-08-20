#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
一次跑完這個資料夾裡所有客戶腳本, 印出每一支的起訖時間與耗時, 全程留一份 log。

為什麼要這支:
  - 那些腳本結尾都有 input("Press Enter to exit...")。直接用 subprocess 跑會卡在
    那裡等人按 Enter, 所以這裡先把換行灌進 stdin 再關掉。
  - 由小到大排, 資料量最小的先跑。萬一有結構性問題, 一分鐘就知道,
    不必等四十分鐘跑完最大那個才發現。
  - 輸出同時印在畫面與寫進 log 檔, 跑完可以回頭查。

**不會修改客戶腳本, 也不會刪任何檔案。**
--clean 只是把上一次的產出「移到」_prev_outputs/ 這個子資料夾, 不是刪除。

用法:
    python run_all.py
    python run_all.py --clean       # 先把上次的產出移走 (見下方 D1 說明)
    python run_all.py --only 2      # 只跑第 2 支 (編號見畫面上的清單)
"""
from __future__ import annotations

import argparse
import ast
import os
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

# ================================================================= 設定
# 腳本清單是「跑的時候自己找」的, 不寫死 —— 寫死等於把客戶的資料夾名
# 抄進這個 public repo, 而且腳本改名就要跟著改。
SELF = {"run_all.py", "survey.py"}
LOG = "run_all.log"
PREV = "_prev_outputs"
# 客戶腳本會產出的檔名樣式 (--clean 用)
OUT_PAT = ["Duplicate_Analysis.xlsx", "Duplicate_Details.csv"]
# ================================================================= 設定結束


def safe_path(p: str) -> str:
    """絕對路徑會帶出客戶名、專案名、Windows 使用者名 —— 這支的輸出會寫進
    run_all.log, 不該留那些。只保留最後一段, 足夠辨認是哪個資料夾。"""
    q = p.replace("\\", "/")
    if q.startswith("/") or (len(q) > 1 and q[1] == ":"):
        return ".../" + q.rstrip("/").rsplit("/", 1)[-1]
    return p


def find_scripts(root: Path) -> list[tuple[str, str, int]]:
    """找出客戶腳本, 並從每支的 folder 常數推出它讀哪個資料夾。

    回傳 (腳本檔名, 資料夾相對路徑, 資料夾位元組) 並依大小由小到大排序 ——
    最小的先跑, 有結構性問題一分鐘就知道, 不必等最大那個跑完四十分鐘才發現。
    """
    out = []
    for p in sorted(root.glob("*.py")):
        if p.name in SELF:
            continue
        folder = ""
        try:
            tree = ast.parse(p.read_text(encoding="utf-8", errors="replace"))
        except SyntaxError:
            out.append((p.name, "(這支語法有問題)", 0))
            continue
        for n in ast.walk(tree):
            if isinstance(n, ast.Assign) and getattr(n.targets[0], "id", "") == "folder":
                try:
                    folder = str(ast.literal_eval(n.value))
                except Exception:
                    pass
        d = (root / folder.replace("\\", "/")) if folder else None
        size = 0
        if not folder:
            folder = "(讀不出 folder)"
        elif d is None or not d.is_dir():
            # 靜靜印一個 "?" 太容易被忽略 —— 這支腳本會 glob 到空清單,
            # 跑完是 0 列而且 exit code 還是 0, 看起來像成功。
            # 找得到正確的資料夾就直接把該寫的那一行給出來 —— 少一個分隔符、
            # 少了 data/ 這一層, 是實際踩過的兩種寫法, 用肉眼對很容易再漏一次。
            fix = ""
            base = folder.replace("\\", "/").lstrip("./")
            for cand in (root / base, root / "data" / base,
                         root / "data" / base.rsplit("/", 1)[-1]):
                if cand.is_dir():
                    rel = cand.relative_to(root).as_posix().replace("/", "\\")
                    fix = f'\n       -> 改成  folder = r".\\{rel}"'
                    break
            folder = f"!! 資料夾不存在: {safe_path(folder)}{fix}"
        else:
            size = sum(f.stat().st_size for f in d.rglob("*") if f.is_file())
        out.append((p.name, folder, size))
    return sorted(out, key=lambda t: t[2])


def mb(n: int) -> str:
    return f"{n / 1e6:,.0f}MB" if n else "?"


def stamp() -> str:
    return datetime.now().strftime("%H:%M:%S")


def hms(sec: float) -> str:
    m, s = divmod(int(sec), 60)
    h, m = divmod(m, 60)
    return f"{h}:{m:02d}:{s:02d}" if h else f"{m}:{s:02d}"


def move_old_outputs(root: Path, scripts, log) -> int:
    """把上一次的產出移走 —— 移動不是刪除, 隨時可以搬回來。

    為什麼要移: 有幾支腳本的排除清單字串大小寫跟它實際的輸出檔名對不上,
    而排除是精確比對, 所以第二次跑會把上一次的 xlsx 當成輸入讀進去。
    詳見 CHANGES.local.md D1。
    """
    dest = root / PREV / datetime.now().strftime("%Y%m%d_%H%M%S")
    n = 0
    for _, folder, _ in scripts:
        d = root / folder.replace("\\", "/")
        if not d.is_dir():
            continue
        for f in d.iterdir():
            if f.is_file() and any(p in f.name for p in OUT_PAT):
                dest.mkdir(parents=True, exist_ok=True)
                f.rename(dest / f.name)
                log(f"    移走 {folder}/{f.name}")
                n += 1
    return n


def run_one(root: Path, script: str, folder: str, size: int,
            i: int, total: int, log) -> dict:
    p = root / script
    log("")
    log("=" * 72)
    log(f"[{i}/{total}] {script}")
    log(f"        讀 {folder}   約 {mb(size)}")
    log(f"        開始 {stamp()}")
    log("=" * 72)
    if not p.exists():
        log(f"  !! 找不到 {script}, 跳過")
        return {"script": script, "rc": None, "secs": 0.0, "note": "檔案不存在"}

    t0 = time.time()
    env = dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUNBUFFERED="1")
    try:
        proc = subprocess.Popen(
            [sys.executable, "-u", str(p)],
            cwd=str(root),
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            env=env,
        )
        # 那三個 input("Press Enter to exit...") 不餵就會卡住
        try:
            proc.stdin.write("\n" * 20)
            proc.stdin.flush()
        except Exception:
            pass
        finally:
            try:
                proc.stdin.close()
            except Exception:
                pass
        for line in proc.stdout:
            log("  " + line.rstrip("\n"))
        proc.wait()
        rc = proc.returncode
    except Exception as e:
        log(f"  !! 啟動失敗: {type(e).__name__}: {e}")
        return {"script": script, "rc": None, "secs": time.time() - t0,
                "note": f"{type(e).__name__}"}

    secs = time.time() - t0
    log("")
    log(f"        結束 {stamp()}   耗時 {hms(secs)}   exit code {rc}")
    return {"script": script, "rc": rc, "secs": secs,
            "note": "正常" if rc == 0 else "非正常結束"}


def main() -> None:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", default=".")
    ap.add_argument("--clean", action="store_true",
                    help=f"先把上次的產出移到 {PREV}/ (移動, 不是刪除)")
    ap.add_argument("--list", action="store_true", dest="list_only",
                    help="只列出執行順序與編號, 不跑任何東西")
    ap.add_argument("--only", type=int, default=None,
                    help="只跑第 N 支 (編號見畫面上的清單)")
    args = ap.parse_args()
    root = Path(args.root).resolve()

    lines: list[str] = []

    def log(s: str = "") -> None:
        lines.append(s)
        print(s, flush=True)

    t0 = time.time()
    log(f"開始 {datetime.now():%Y-%m-%d %H:%M:%S}")
    log(f"工作目錄 {root}")
    log(f"Python {sys.version.split()[0]}")
    scripts = find_scripts(root)
    if not scripts:
        sys.exit("這個資料夾裡找不到客戶腳本 (*.py)")
    log("")
    log("執行順序 (依資料量由小到大, 有問題早點發現):")
    bad = []
    for i, (nm, fd, sz) in enumerate(scripts, 1):
        log(f"  {i}. {nm:<50} {mb(sz):>8}   {fd if sz == 0 else ''}".rstrip())
        if fd.startswith("!!") or fd.startswith("("):
            bad.append((i, nm, fd))
    if bad:
        log("")
        log("!! 下面這幾支讀不到資料, 跑了也是 0 列 (而且 exit code 仍是 0):")
        for i, nm, fd in bad:
            log(f"     {i}. {nm}  ->  {fd}")
        log("   路徑對不上就先修 folder 那一行, 不要直接跑。")

    # 編號會隨資料夾大小變動 (修好一支的路徑, 它的排序就跟著變),
    # 所以 --only 之前先用 --list 確認一次。
    if args.list_only:
        log("")
        log("(--list: 只列清單, 沒有執行任何腳本)")
        return

    # 先決定要跑哪幾支, --clean 才知道該清誰 —— 以前是無條件清掉**所有**
    # 資料夾的產出, 配上 --only 就等於把其他三支剛跑好的結果丟進 _prev_outputs/。
    todo = list(enumerate(scripts, 1))
    if args.only:
        todo = [t for t in todo if t[0] == args.only]
        if not todo:
            sys.exit(f"--only 要在 1..{len(scripts)} 之間")
    else:
        # 讀不到資料的不自動跑 —— 資料夾裡所有 *.py 都會被掃進來,
        # 其中可能有根本不是客戶 combine 腳本的東西; 路徑寫錯的也一樣,
        # 跑了只會產出 0 列而 exit code 還是 0。真要跑就 --only 指名。
        skipped = [t for t in todo if t[1][1].startswith(("!!", "("))]
        todo = [t for t in todo if not t[1][1].startswith(("!!", "("))]
        for i, (nm, fd, _) in skipped:
            log(f"  跳過 [{i}] {nm} —— {fd}  (要跑就用 --only {i})")

    if args.clean:
        log("")
        log(f"--clean: 把**這次要跑的**腳本的舊產出移到 {PREV}/")
        n = move_old_outputs(root, [t[1] for t in todo], log)
        log(f"    共移走 {n} 個檔" if n else "    沒有找到上次的產出")
    else:
        log("")
        log("(沒有加 --clean。若資料夾裡還留著上次的產出, 其中三支會把它當成")
        log(" 輸入讀進去 —— 每跑一次 Total Rows 就多一列並累積下去,")
        log(" 但不會改變 keep/adjust 判斷。詳見 CHANGES.local.md L7。)")

    results = []
    for i, (script, folder, size) in todo:
        results.append(run_one(root, script, folder, size, i, len(scripts), log))

    total = time.time() - t0
    log("")
    log("=" * 72)
    log("總結")
    log("=" * 72)
    log(f"  {'腳本':<52}{'耗時':>10}{'結果':>10}")
    for r in results:
        log(f"  {r['script']:<52}{hms(r['secs']):>10}{r['note']:>10}")
    log(f"  {'總計':<52}{hms(total):>10}")
    bad = [r for r in results if r["rc"] != 0]
    log("")
    if bad:
        log(f"  !! 有 {len(bad)} 支沒有正常結束:")
        for r in bad:
            log(f"       {r['script']}  ({r['note']})")
        log("     往上捲看那一支的輸出, 錯誤訊息在它的區塊裡。")
    else:
        log(f"  {len(results)} 支都正常結束。")
    log("")
    log(f"結束 {datetime.now():%Y-%m-%d %H:%M:%S}")

    out = root / LOG
    out.write_text("\n".join(lines), encoding="utf-8")
    print(f"\nlog: {out}")


if __name__ == "__main__":
    main()
