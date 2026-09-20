#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
duck_shim.py —— 讓一支用 sqlite3 寫的腳本改用 DuckDB 當引擎, 腳本本身只改一行。

    import duck_shim
    conn = duck_shim.connect(db_path)      # 原本是 sqlite3.connect(db_path)

為什麼要這樣做: 那支腳本有四千多行、18 處 pd.read_sql、12 處 executescript、
17 處 PRAGMA。散在裡面逐處改, 每一處都是自己引入 bug 的機會, 而且 re-performance
的說法會變弱 —— 改得越少, 「我們重跑的是客戶的邏輯」這句話越站得住。所以把差異
全部收在這一層, 客戶的 SQL 一句不動。

實測 (duckdb 1.5, 2026-09-17) 兩邊行為的差異與處理:

  PRAGMA journal_mode / synchronous / cache_size ...
      DuckDB 不認識, 會丟 CatalogException。這些是 sqlite 的效能旋鈕, 不參與
      計算 —— 換成一句空的 SELECT。注意不能「什麼都不做」: 腳本有
          pd.read_sql_query("PRAGMA quick_check", conn).iloc[0, 0]
      這種寫法, 沒有結果集的話 pandas 讀 cursor.description 會拿到 None。

  PRAGMA quick_check / integrity_check
      sqlite 用來確認檔案沒壞。DuckDB 沒有對應的語法, 但可以做等效的事:
      CHECKPOINT 會把資料強制寫出並重讀目錄, 檔案有問題就會丟例外。成功才
      回 'ok', 失敗照實回報 —— 不是無條件回 'ok'。

  ORDER BY 遇到 NULL
      sqlite 把 NULL 當最小值 (ASC 排最前、DESC 排最後); DuckDB 預設兩邊都排
      最後。腳本用 ORDER BY 加視窗函數決定同一組重複裡哪一列標 Keep, 排序差一
      點結果就不同。連線時設 default_null_order 讓它跟 sqlite 一致。

  executescript()
      DuckDB 沒有這個方法, 但它的 execute() 本來就吃得下用分號隔開的多句。
      直接轉給 execute。

  INSERT OR REPLACE / OR IGNORE
      DuckDB 要求目標表上有 UNIQUE 或 PRIMARY KEY, 否則報錯。而 sqlite 在沒有
      約束時, INSERT OR REPLACE 的行為跟普通 INSERT 完全一樣。所以先查目標表
      有沒有約束: 沒有就改寫成普通 INSERT (等價, 並記一筆); 有就原樣丟過去讓
      DuckDB 自己處理。

  WITHOUT ROWID
      sqlite 的建表儲存選項 (改用主鍵當儲存鍵, 不要隱含的 rowid)。DuckDB 的
      parser 直接報錯。這是儲存配置不是邏輯, 查詢結果一模一樣 —— 拿掉。

  GROUP_CONCAT 的元素順序   ** 已知差異, 這一層改不了 **
      沒有 ORDER BY 的 GROUP_CONCAT, 兩個引擎的元素順序都是未定義的 —— sqlite
      的順序跟著查詢計畫走 (有沒有用到索引都會變), DuckDB 有自己的一套。實測
      內容完全相同, 只有排列不同:
          sqlite  1|1|1|1|3|3|3
          duckdb  1|3|1|3|1|3|1
      要在補一個 ORDER BY 才會一致, 但那等於改客戶的 SQL。所以不改, 改成比對
      的時候把 "|" 串起來的多值當成無序集合 (csvtools.py --cmp --loose)。

  文字欄與數字比較
      sqlite 的 '007' = 7 是 false (不隱式轉型), DuckDB 是 true。這一層管不到,
      要人去確認 SQL 裡有沒有這種寫法。

pandas 的 read_sql_query 直接吃 DuckDB 連線就對, 那 18 處不用動。

to_sql 不行 —— 能跑, 但慢得離譜。pandas 走 DBAPI 路徑時會把 DataFrame 拆成
一句句 INSERT ... VALUES (?,?,...) 再 executemany; DuckDB 的單筆 INSERT 遠比
sqlite 慢, 而這個呼叫就在逐塊載入的迴圈裡。實測 5 萬列 x 54 欄:

    pandas to_sql (DBAPI 逐筆)        25.10s
    register + CREATE AS SELECT        0.25s     快 100x
    INSERT INTO ... BY NAME SELECT     1.21s     快  21x

換算三千萬列是「四小時」對「十幾分鐘」。所以這裡把 DataFrame.to_sql 換掉:
連線是本模組的就用 DuckDB 原生的 register + SQL, 其他連線原樣交還給 pandas。
資料完全一樣, 只是搬運方式不同。
"""
from __future__ import annotations

import os
import re
import warnings

import duckdb

# sqlite 的 NULL 當最小值: ASC 排最前、DESC 排最後
NULL_ORDER = "nulls_first_on_asc_last_on_desc"

_PRAGMA = re.compile(r"^\s*PRAGMA\b", re.I)
_LEAD = re.compile(r"^(?:\s+|--[^\n]*\n?|/\*.*?\*/)+", re.S)
_NOROWID = re.compile(r"\)\s*WITHOUT\s+ROWID", re.I)
_INTEGRITY = re.compile(
    r"^\s*PRAGMA\s+\w*(quick_check|integrity_check|foreign_key_check)", re.I)
_PRAGMA_NAME = re.compile(r"^\s*PRAGMA\s+([\w.]+)", re.I)
_INSERT_OR = re.compile(
    r"^(\s*)INSERT\s+OR\s+(REPLACE|IGNORE|ROLLBACK|ABORT|FAIL)\s+INTO\s+"
    r"[\"'`\[]?(\w+)", re.I)

# DuckDB 自己就有這幾個, 原樣送過去比自己編一個假的好
_PRAGMA_NATIVE = {"table_info", "show_tables", "database_list",
                  "database_size", "storage_info", "version", "collations",
                  "functions", "show_tables_expanded"}
# 腳本拿來報工作檔多大的。sqlite 是 page_count * page_size, DuckDB 沒有
# 分頁的概念, 就用檔案實際大小換算回去, 讓那個數字仍然是真的。
_PRAGMA_SIZE = {"page_count", "page_size", "freelist_count"}
_PAGE = 4096

# GLOB 的字元類別否定, 兩邊寫法不同:
#   sqlite  [^0-9]  = 不是數字
#   DuckDB  [!0-9]  = 不是數字;  DuckDB 把 [^0-9] 讀成「^ 或數字」
# 所以 '*[^0-9]*' 直接送過去意思會**剛好相反**, 而且不會報錯, 只會安靜地
# 篩出錯的列。腳本用它判斷航班號是不是純數字, 屬於計算邏輯, 必須改對。
# 另外 DuckDB 的語法裡沒有 NOT GLOB, 只有 GLOB。改寫成 `GLOB p = FALSE`:
# 三值邏輯下 (A = FALSE) 與 NOT A 完全等價 —— TRUE->FALSE、FALSE->TRUE、
# NULL->NULL, 而且 = 比 AND/OR 先結合, 不必去解析左邊那串運算式。
_GLOB_LIT = re.compile(r"\bGLOB\s*('(?:[^']|'')*')", re.I)
_GLOB_NOT = re.compile(r"\bNOT\s+GLOB\s*('(?:[^']|'')*')", re.I)
_GLOB_ANY = re.compile(r"\bGLOB\b", re.I)

# sqlite 的 LIKE 對 ASCII 大小寫不敏感, DuckDB 的 LIKE 敏感。腳本靠
# LIKE 'Keep%' / 'Adjust%' 判斷處理方式, 換成 ILIKE 才是同一個意思。
_LIKE_ANY = re.compile(r"\bLIKE\b", re.I)
_LIKE_SUB = re.compile(r"\bLIKE\b", re.I)

# sqlite 的 CAST(文字 AS INTEGER) 取開頭那段數字, 取不到就 0, 不會出錯;
# DuckDB 嚴格, 'CX043' 直接丟 Conversion Error。腳本在 GLOB 判斷之後才
# CAST, 但 SQL 不保證 AND/OR/CASE 的求值順序, 空字串也會通過那個判斷。
_CAST_ANY = re.compile(r"\bCAST\s*\(", re.I)
_CAST_INT = {"INTEGER", "INT", "BIGINT", "SMALLINT", "TINYINT",
             "INT2", "INT4", "INT8", "MEDIUMINT"}
_CAST_EMU = ("CASE WHEN ({e}) IS NULL THEN NULL ELSE COALESCE(TRY_CAST("
             "regexp_extract(CAST(({e}) AS VARCHAR), '^\\s*[-+]?[0-9]+')"
             " AS BIGINT), 0) END")

# sqlite 的 ROUND 是把 double 的**精確**十進位展開拿去四捨五入(逢五進位),
# DuckDB 會先當成十進位數, 於是 2.675 一個給 2.67 一個給 2.68。printf 拿得
# 到精確展開, 再交給 DECIMAL 進位就跟 sqlite 一致。
# sqlite 的 REAL 是 8 bytes; DuckDB 的 REAL/FLOAT 是 4 bytes。金額欄位宣告成
# REAL 的話, 2759.59 存進去會變 2759.590087890625, 而且不會有任何提示。
_REAL_ANY = re.compile(r"\b(?:REAL|FLOAT4?|NUMERIC|DECIMAL)\b(?!\s*\()", re.I)

# 索引跟 PRAGMA 一樣是效能結構, 不參與計算。sqlite 是列式儲存, 沒索引就得
# 全表掃描, 所以腳本建了一堆; DuckDB 是欄式加 zone map, 這些索引幫不上忙,
# 建起來反而吃記憶體跟時間。而且 DuckDB 不支援部分索引 (CREATE INDEX ...
# WHERE ...)。非 UNIQUE 的索引整個跳過。
# UNIQUE 索引不一樣 —— 它會擋重複列, 有語意, 要保留。
_CREATE_IDX = re.compile(r"^\s*CREATE\s+(UNIQUE\s+)?INDEX\b", re.I)
_IDX_NAME = re.compile(
    r"^\s*CREATE\s+(?:UNIQUE\s+)?INDEX\s+(?:IF\s+NOT\s+EXISTS\s+)?"
    r"[\"'`\[]?([\w.]+)", re.I)
_DROP_IDX = re.compile(r"^\s*DROP\s+INDEX\s+(?!IF\s+EXISTS)", re.I)

# sqlite 可以把索引建在指定的 schema 裡 (CREATE INDEX temp.idx ON ...),
# DuckDB 的語法不吃索引名前面的 schema。索引名只是個標籤, 拿掉限定詞不改
# 變它管哪張表, 更不改變 UNIQUE 擋不擋重複列。
_IDX_QUAL = re.compile(
    r"^(\s*CREATE\s+(?:UNIQUE\s+)?INDEX\s+(?:IF\s+NOT\s+EXISTS\s+)?)"
    r"([\w\"'`\[\]]+)\.(\w+)", re.I)
_DROPIDX_QUAL = re.compile(
    r"^(\s*DROP\s+INDEX\s+(?:IF\s+EXISTS\s+)?)([\w\"'`\[\]]+)\.(\w+)", re.I)

_ROUND_ANY = re.compile(r"\bROUND\s*\(", re.I)
_ROUND_EMU = ("CAST(ROUND(CAST(printf('%.25f', CAST(({e}) AS DOUBLE))"
              " AS DECIMAL(38,25)), {n}) AS DOUBLE)")

notes: list[str] = []          # 這一層實際做過什麼, 可以抄進 CHANGES 記錄


def _note(msg: str) -> None:
    if msg not in notes:
        notes.append(msg)
        print(f"  [duck_shim] {msg}", flush=True)


class _Wrapped:
    """把 PRAGMA 與 INSERT OR 攔下來之後再轉給 DuckDB。"""

    def __init__(self, inner, owner):
        self._inner = inner
        self._owner = owner

    # --- 轉譯 ---------------------------------------------------------
    def _prep(self, sql):
        # 前導的空白與註解要先跳過再認關鍵字。客戶腳本寫的是
        #     /* 效能設定 */ PRAGMA auto_vacuum = NONE;
        # 直接用 ^\s*PRAGMA 對不上, PRAGMA 就漏出去給 DuckDB 了。
        m0 = _LEAD.match(sql)
        off = m0.end() if m0 else 0
        head = sql[off:]
        if _PRAGMA.match(head):
            if _INTEGRITY.match(head):
                # 做一次等效的檢查再回答, 不要無條件說 ok。CHECKPOINT 會
                # 把資料寫出並重讀目錄, 檔案有問題就會丟例外。
                try:
                    self._inner.execute("CHECKPOINT")
                except Exception as e:
                    _note(f"!! CHECKPOINT 失敗: {str(e).splitlines()[0][:60]}")
                    return ("SELECT 'checkpoint failed' AS integrity_check")
                _note("PRAGMA quick_check -> 改用 DuckDB 的 CHECKPOINT, 通過")
                return "SELECT 'ok' AS integrity_check"
            nm = _PRAGMA_NAME.match(head)
            name = (nm.group(1) if nm else "?").lower()
            if name in _PRAGMA_NATIVE:
                return sql
            if name in _PRAGMA_SIZE:
                val = self._owner._size_pragma(name)
                _note(f"PRAGMA {name} -> "
                      + ("回 0 (DuckDB 沒有 freelist)"
                         if name == "freelist_count"
                         else "改用 DuckDB 檔案的實際大小換算"))
                return f'SELECT {val} AS "{name}"'
            # 帶括號的是表格型 (index_list(t) 之類), 回空結果集;
            # 不帶括號的是純量, 一定要回得出一列, 否則 fetchone()[0] 會炸。
            rest = head[nm.end():].lstrip() if nm else ""
            _note(f"PRAGMA {name} 略過 (sqlite 的效能旋鈕, 不參與計算)")
            if rest.startswith("("):
                return "SELECT NULL AS pragma_ignored WHERE FALSE"
            return f'SELECT 0 AS "{name}"'
        m = _CREATE_IDX.match(head)
        if m:
            nm = _IDX_NAME.match(head)
            nm = nm.group(1) if nm else "?"
            if not m.group(1):                 # 不是 UNIQUE, 純效能
                _note(f"CREATE INDEX {nm} 跳過 (索引不參與計算; "
                      "DuckDB 是欄式儲存, 不需要)")
                return "SELECT NULL AS index_skipped WHERE FALSE"
            if re.search(r"\bWHERE\b", head, re.I):
                raise RuntimeError(
                    f"duck_shim: {nm} 是部分 UNIQUE 索引, DuckDB 不支援, "
                    "而且它會影響哪些列算重複, 不能跳過。\n  " +
                    " ".join(head.split())[:300])
            _note(f"CREATE UNIQUE INDEX {nm} 保留 (會擋重複列, 有語意)")
            q = _IDX_QUAL.match(head)
            if q:
                _note(f"索引名 {q.group(2)}.{q.group(3)} -> {q.group(3)} "
                      "(DuckDB 的語法不吃索引名前面的 schema)")
                head = _IDX_QUAL.sub(r"\1\3", head, count=1)
                sql = sql[:off] + head
        if _DROP_IDX.match(head) or _DROPIDX_QUAL.match(head):
            if _DROP_IDX.match(head):
                head = _DROP_IDX.sub(
                    lambda x: x.group(0).rstrip() + " IF EXISTS ", head,
                    count=1)
            head = _DROPIDX_QUAL.sub(r"\1\3", head, count=1)
            sql = sql[:off] + head
        if _GLOB_ANY.search(sql):
            sql = _fix_glob(sql)
        if _LIKE_ANY.search(sql):
            sql = _fix_like(sql)
        if _REAL_ANY.search(sql):
            sql = _fix_real(sql)
        if _CAST_ANY.search(sql):
            sql = _fix_call(sql, "CAST", _cast_emu)
        if _ROUND_ANY.search(sql):
            sql = _fix_call(sql, "ROUND", _round_emu)
        head = sql[off:]
        if _NOROWID.search(sql):
            sql = _NOROWID.sub(")", sql)
            head = sql[off:]
            _note("拿掉 WITHOUT ROWID (sqlite 的儲存配置, 不影響查詢結果)")
        m = _INSERT_OR.match(head)
        if m:
            table = m.group(3)
            if not self._owner._has_unique(table):
                # 沒有 UNIQUE/PK 時, sqlite 的 INSERT OR xxx 等同普通 INSERT
                _note(f"{table} 沒有 UNIQUE/PK 約束, "
                      f"INSERT OR {m.group(2).upper()} 視為普通 INSERT (等價)")
                return sql[:off] + _INSERT_OR.sub(r"\1INSERT INTO \3",
                                                  head, count=1)
            _note(f"{table} 有約束, INSERT OR {m.group(2).upper()} 原樣交給 DuckDB")
        return sql

    # --- DB-API ------------------------------------------------------
    def execute(self, sql, params=()):
        sql = self._prep(sql)
        if sql is None:
            return self
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            if params:
                self._inner.execute(sql, list(params))
            else:
                self._inner.execute(sql)
        return self

    def executescript(self, sql):
        """sqlite 專屬。DuckDB 的 execute 本來就吃多句, 但 PRAGMA 要先濾掉,
        所以按分號拆開逐句送 —— 拆的時候要避開字串常值裡的分號。"""
        for stmt in _split(sql):
            if stmt.strip():
                self.execute(stmt)
        return self

    def executemany(self, sql, seq):
        sql = self._prep(sql)
        if sql is not None:
            self._inner.executemany(sql, [list(x) for x in seq])
        return self

    def fetchall(self):
        return self._inner.fetchall()

    def fetchone(self):
        return self._inner.fetchone()

    def fetchmany(self, size=1):
        return self._inner.fetchmany(size)

    @property
    def description(self):
        return self._inner.description

    def cursor(self):
        return _Wrapped(self._inner.cursor(), self._owner)

    def commit(self):
        return None                            # DuckDB 自動提交

    def rollback(self):
        return None

    def close(self):
        return self._inner.close()

    def __enter__(self):
        return self

    def __exit__(self, *a):
        return False

    def __getattr__(self, name):               # df()、fetchdf() 之類的原生方法
        return getattr(self._inner, name)


class Connection(_Wrapped):
    def __init__(self, path):
        inner = duckdb.connect(str(path))
        inner.execute(f"SET default_null_order='{NULL_ORDER}'")
        _Wrapped.__init__(self, inner, self)
        self._uniq: dict = {}
        self._path = str(path)

    def _size_pragma(self, name: str) -> int:
        """page_count * page_size 要等於檔案真正佔的空間。不呼叫 CHECKPOINT
        —— 那是實際的寫入工作, 腳本每載一批就問一次大小, 會拖慢整輪。
        未 checkpoint 的部分在 .wal 裡, 一起算進去才是真的磁碟用量。"""
        if name == "page_size":
            return _PAGE
        if name == "freelist_count":
            return 0
        total = 0
        for p in (self._path, self._path + ".wal"):
            try:
                total += os.path.getsize(p)
            except OSError:
                pass
        return total // _PAGE

    def _table_exists(self, name: str) -> bool:
        try:
            self._inner.execute(f'SELECT 1 FROM "{name}" LIMIT 0')
            return True
        except Exception:
            return False

    def _has_unique(self, table: str) -> bool:
        t = table.lower()
        if t not in self._uniq:
            try:
                rows = self._inner.execute(
                    "SELECT count(*) FROM duckdb_constraints() "
                    "WHERE lower(table_name) = ? "
                    "AND constraint_type IN ('PRIMARY KEY','UNIQUE')",
                    [t]).fetchone()
                self._uniq[t] = bool(rows and rows[0])
            except Exception:
                self._uniq[t] = False
        return self._uniq[t]


def _fix_glob(sql: str) -> str:
    """把 GLOB 後面那個字面值裡的 [^ 改成 [!。只動字面值本身, 不碰其他。

    對不上字面值的 GLOB (例如 pattern 是欄位或參數) 會直接丟例外 —— 這種
    情形沒辦法在字串層面改對, 而錯的結果不會報錯, 寧可停在這裡。"""
    if len(_GLOB_LIT.findall(sql)) != len(_GLOB_ANY.findall(sql)):
        raise RuntimeError(
            "duck_shim: 有 GLOB 後面不是字串常值, 無法確認語意是否相同。\n"
            "  sqlite 的 [^...] 與 DuckDB 的 [^...] 意思相反, 不能就這樣送過去。\n"
            "  出問題的 SQL:\n    " + " ".join(sql.split())[:400])
    changed = []

    def fix(lit):
        new = lit.replace("[^", "[!")
        if new != lit:
            changed.append((lit, new))
        return new

    out = _GLOB_NOT.sub(lambda m: f"GLOB {fix(m.group(1))} = FALSE", sql)
    if out != sql:
        _note("NOT GLOB p -> GLOB p = FALSE "
              "(DuckDB 沒有 NOT GLOB; 三值邏輯下兩者等價)")
    out = _GLOB_LIT.sub(
        lambda m: m.group(0)[:-len(m.group(1))] + fix(m.group(1)), out)
    for old, new in dict.fromkeys(changed):
        _note(f"GLOB {old} -> {new} (字元類別否定, sqlite 用 ^ / DuckDB 用 !)")
    return out


def _lit_spans(sql: str) -> list:
    """字串常值與註解的範圍。改寫只能動這些範圍**以外**的字。"""
    spans, i, n = [], 0, len(sql)
    while i < n:
        c = sql[i]
        if c in "'\"":
            j = i + 1
            while j < n:
                if sql[j] == c:
                    if j + 1 < n and sql[j + 1] == c:
                        j += 2
                        continue
                    break
                j += 1
            spans.append((i, min(j + 1, n)))
            i = j + 1
        elif c == "-" and i + 1 < n and sql[i + 1] == "-":
            j = sql.find("\n", i)
            j = n if j < 0 else j
            spans.append((i, j))
            i = j
        elif c == "/" and i + 1 < n and sql[i + 1] == "*":
            j = sql.find("*/", i + 2)
            j = n if j < 0 else j + 2
            spans.append((i, j))
            i = j
        else:
            i += 1
    return spans


def _free(pos: int, spans: list) -> bool:
    return not any(a <= pos < b for a, b in spans)


def _calls(sql: str, name: str) -> list:
    """由後往前列出 NAME(...) 的 (起, 迄, 括號內容)。由後往前才不會改壞位置。"""
    spans = _lit_spans(sql)
    out = []
    for m in re.finditer(r"\b" + name + r"\s*\(", sql, re.I):
        if not _free(m.start(), spans):
            continue
        depth, j = 1, m.end()
        while j < len(sql) and depth:
            if _free(j, spans):
                if sql[j] == "(":
                    depth += 1
                elif sql[j] == ")":
                    depth -= 1
            j += 1
        if depth == 0:
            out.append((m.start(), j, sql[m.end():j - 1]))
    return out[::-1]


def _top_split(inner: str, sep: str = ","):
    """最外層的逗號才算分隔, 括號裡與常值裡的不算。"""
    spans = _lit_spans(inner)
    parts, depth, last = [], 0, 0
    for i, c in enumerate(inner):
        if not _free(i, spans):
            continue
        if c == "(":
            depth += 1
        elif c == ")":
            depth -= 1
        elif c == sep and depth == 0:
            parts.append(inner[last:i])
            last = i + 1
    parts.append(inner[last:])
    return parts


def _cast_emu(inner: str):
    """CAST(x AS INTEGER) -> 跟 sqlite 同語意的運算式; 其他型別不動。"""
    spans = _lit_spans(inner)
    depth, cut = 0, None
    for i, c in enumerate(inner):
        if not _free(i, spans):
            continue
        if c == "(":
            depth += 1
        elif c == ")":
            depth -= 1
        elif depth == 0 and c.upper() == "A" and inner[i:i + 2].upper() == "AS":
            before = inner[i - 1] if i else " "
            after = inner[i + 2] if i + 2 < len(inner) else " "
            if not (before.isalnum() or before == "_") and \
               not (after.isalnum() or after == "_"):
                cut = i
    if cut is None:
        return None
    if inner[cut + 2:].strip().upper() not in _CAST_INT:
        return None
    expr = inner[:cut].strip()
    _note("CAST(x AS INTEGER) -> 取開頭數字、取不到給 0 "
          "(sqlite 的轉型規則; DuckDB 原本會直接報錯)")
    return _CAST_EMU.format(e=expr)


def _round_emu(inner: str):
    parts = _top_split(inner)
    if len(parts) == 1:
        expr, nd = parts[0].strip(), "0"
    elif len(parts) == 2:
        expr, nd = parts[0].strip(), parts[1].strip()
    else:
        return None
    _note("ROUND -> 改用 double 的精確十進位展開再進位 (sqlite 的作法)")
    return _ROUND_EMU.format(e=expr, n=nd)


def _fix_call(sql: str, name: str, emu, depth: int = 0) -> str:
    """只處理最外層的 NAME(...), 裡面的先遞迴處理好再包進去。

    改出來的字串不會再被掃一次 —— ROUND 的替代寫法裡自己就有 ROUND,
    重掃會無限循環。"""
    if depth > 20:
        raise RuntimeError(f"duck_shim: {name}() 巢狀太深, 不敢改")
    tops = []
    for s, e, inner in _calls(sql, name)[::-1]:        # 由前往後
        if tops and s < tops[-1][1]:                   # 被前一個包住
            continue
        tops.append((s, e, inner))
    for s, e, inner in reversed(tops):                 # 由後往前才不會改壞位置
        inner2 = _fix_call(inner, name, emu, depth + 1)
        new = emu(inner2)
        if new is None:
            head = len(sql[s:e]) - len(inner) - 1
            new = sql[s:s + head] + inner2 + ")"
        sql = sql[:s] + new + sql[e:]
    return sql


def _sub_free(sql: str, pat, repl: str) -> str:
    """只換字串常值與註解**以外**的部分。"""
    spans = _lit_spans(sql)
    out, last, n = [], 0, 0
    for m in pat.finditer(sql):
        if not _free(m.start(), spans):
            continue
        out.append(sql[last:m.start()])
        out.append(repl)
        last = m.end()
        n += 1
    if not n:
        return sql
    out.append(sql[last:])
    return "".join(out)


def _fix_like(sql: str) -> str:
    out = _sub_free(sql, _LIKE_SUB, "ILIKE")
    if out != sql:
        _note("LIKE -> ILIKE (sqlite 的 LIKE 對 ASCII 大小寫不敏感)")
    return out


def _fix_real(sql: str) -> str:
    out = _sub_free(sql, _REAL_ANY, "DOUBLE")
    if out != sql:
        _note("REAL/FLOAT/NUMERIC -> DOUBLE "
              "(sqlite 的 REAL 是 8 bytes, DuckDB 的只有 4 bytes)")
    return out


def _split(script: str) -> list:
    """按分號拆成多句, 但字串常值與註解裡的分號不算。"""
    out, buf = [], []
    i, n = 0, len(script)
    quote = None
    while i < n:
        c = script[i]
        if quote:
            buf.append(c)
            if c == quote:
                if i + 1 < n and script[i + 1] == quote:   # '' 跳脫
                    buf.append(script[i + 1])
                    i += 1
                else:
                    quote = None
        elif c in "'\"":
            quote = c
            buf.append(c)
        elif c == "-" and i + 1 < n and script[i + 1] == "-":
            while i < n and script[i] != "\n":
                buf.append(script[i])
                i += 1
            continue
        elif c == "/" and i + 1 < n and script[i + 1] == "*":
            j = script.find("*/", i + 2)
            j = n if j < 0 else j + 2
            buf.append(script[i:j])          # 註解留著, 但裡面的分號不算句尾
            i = j
            continue
        elif c == ";":
            out.append("".join(buf))
            buf = []
        else:
            buf.append(c)
        i += 1
    if "".join(buf).strip():
        out.append("".join(buf))
    return out


# --- pandas 那一側 --------------------------------------------------------
_SQL_TYPE = {"TEXT": "VARCHAR", "VARCHAR": "VARCHAR", "CHAR": "VARCHAR",
             "NVARCHAR": "VARCHAR", "STRING": "VARCHAR",
             "INTEGER": "BIGINT", "INT": "BIGINT", "BIGINT": "BIGINT",
             "REAL": "DOUBLE", "FLOAT": "DOUBLE", "DOUBLE": "DOUBLE",
             "NUMERIC": "DOUBLE", "DECIMAL": "DOUBLE",
             "DATE": "DATE", "TIMESTAMP": "TIMESTAMP", "BLOB": "BLOB"}


def _duck_type(spec) -> str:
    t = str(spec).upper()
    for k, v in _SQL_TYPE.items():
        if t.startswith(k):
            return v
    return "VARCHAR"


def _to_sql(self, name, con, schema=None, if_exists="fail", index=True,
            index_label=None, chunksize=None, dtype=None, method=None, **kw):
    if not isinstance(con, Connection):
        return _ORIG_TO_SQL(self, name, con, schema=schema, if_exists=if_exists,
                            index=index, index_label=index_label,
                            chunksize=chunksize, dtype=dtype, method=method,
                            **kw)
    df = self.reset_index() if index else self
    inner = con._inner
    key = "__duck_shim_df"
    exists = con._table_exists(name)
    if exists and if_exists == "fail":
        raise ValueError(f"Table '{name}' already exists.")
    # 腳本如果指定了欄位型別就照著轉。不轉的話型別由 DataFrame 推, 本來是
    # TEXT 的欄位可能變成 DOUBLE, 比較與排序的語意就跟著變了。
    if dtype:
        cols = ", ".join(
            f'CAST("{c}" AS {_duck_type(dtype[c])}) AS "{c}"' if c in dtype
            else f'"{c}"' for c in df.columns)
    else:
        cols = "*"
    inner.register(key, df)
    try:
        if exists and if_exists == "replace":
            inner.execute(f'DROP TABLE IF EXISTS "{name}"')
            exists = False
        if exists:
            inner.execute(f'INSERT INTO "{name}" BY NAME '
                          f'SELECT {cols} FROM "{key}"')
        else:
            inner.execute(f'CREATE TABLE "{name}" AS SELECT {cols} FROM "{key}"')
    finally:
        try:
            inner.unregister(key)
        except Exception:
            pass
    return len(df)


try:
    import pandas as _pd
    _ORIG_TO_SQL = _pd.DataFrame.to_sql
    _pd.DataFrame.to_sql = _to_sql
    # pandas 只是說「沒有正式測過 DuckDB 連線」。read_sql_query 那條路實測
    # 結果與 sqlite 相同, 所以這句警告沒有資訊量, 關掉。
    warnings.filterwarnings(
        "ignore", message=".*only supports SQLAlchemy connectable.*")
except ImportError:      # 沒裝 pandas 也能用, 只是 to_sql 用不到
    _ORIG_TO_SQL = None


def connect(path, *a, **k):
    """取代 sqlite3.connect。多餘的參數 (timeout、check_same_thread 之類) 忽略。"""
    return Connection(path)


# 讓 `import duck_shim as sqlite3` 可以整組替換掉 sqlite3 —— 腳本只要動 import
# 那一行, connect 的呼叫(可能跨好幾行、帶一堆參數)完全不用碰。底下這幾個名字
# 是 sqlite3 模組有、腳本可能用到的。
Error = duckdb.Error
DatabaseError = duckdb.DatabaseError
OperationalError = duckdb.OperationalError
IntegrityError = duckdb.IntegrityError
ProgrammingError = getattr(duckdb, "ProgrammingError", duckdb.Error)
InterfaceError = getattr(duckdb, "InterfaceError", duckdb.Error)
DataError = getattr(duckdb, "DataError", duckdb.Error)
NotSupportedError = getattr(duckdb, "NotSupportedError", duckdb.Error)
Warning = getattr(duckdb, "Warning", UserWarning)
Cursor = _Wrapped
Row = tuple
PARSE_DECLTYPES = 1
PARSE_COLNAMES = 2
version = duckdb.__version__
sqlite_version = f"duckdb {duckdb.__version__} (via duck_shim)"


def register_adapter(*a, **k):      # sqlite 專屬, DuckDB 用不到
    return None


def register_converter(*a, **k):
    return None
