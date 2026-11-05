# -*- coding: utf-8 -*-
"""
账套 → 业务库 解析器（本项目最关键的一步）
============================================================

【模块职责】
    给定"账套号 + 会计年度"，从用友系统库 UFSystem 的
    UA_AccountDatabase 表里回答**两个问题**：

        ① 本年度数据在哪个业务库？      → AccountTarget.database
        ② 前一年度数据在哪个业务库？    → AccountTarget.prev_database
           （部分报表需要与上年同期对比，如 LK02 利润表"上年同期累计数"）

【为什么这一步必须单独做一个模块？】
    因为它是"数据对不对"的总开关。
    只要这一步错了，后面 43 张报表算得再准也是错的账 —— 而且
    程序**不会报错**，会安安静静地输出另一家公司的数字。
    这种错误比崩溃危险一百倍：崩溃你立刻知道，错账你可能一个月后才发现。

    所以它被单独拆出来，做到三件事：
        ① 规则只有一份（就在这里），不散落在各处；
        ② 出错时给出**人能看懂、能照着处理**的提示，而不是 SQL 报错；
        ③ 结果在切换数据库前还要再验证一次"库真的存在"。

【★ 判定规则（20261105 现场口径，务必按此理解）★】

    设该账套在 UA_AccountDatabase 中共有 N 条记录，
    每条记录有 cDatabase（库名）与 iBeginYear（起始年度），
    按 iBeginYear **升序**排列。

    ── 情形一：N == 1（跨年度单库 / 或只有一年的分年度库）──
        无论查哪一年、也无论查前一年，都只能连这**唯一**的库。
        例：唯一记录 iBeginYear=2013, cDatabase='UFDATA_001'
            查 2026 → UFDATA_001；查前一年 2025 → 也是 UFDATA_001。

        （U8 推荐的"跨年度单库"就是这个形态：一条记录覆盖至今所有年度。）

    ── 情形二：N > 1，且查询年度**恰好等于**某条记录的 iBeginYear ──
        该条记录的 cDatabase 就是本年度业务库。

        前一年库的判定：
          · 若该条 iBeginYear 是**所有记录中最小的**
            → 这一年是首年，**没有前一年数据**（prev_database = None）
          · 否则 → 前一年库取"iBeginYear 数值相邻但较小的那条"的 cDatabase

        例：记录 2013→A、2015→B、2018→C
            查 2015 → 本年库 B；前一年库取 2013 那条 → A
            查 2013 → 本年库 A；2013 是最小，故 prev = None

    ── 情形三：N > 1，且查询年度与所有 iBeginYear 都不相等 ──
        查询年度落在两条相邻记录之间（或大于最大的那条），
        则**本年度与前一年度都使用"iBeginYear 较小的那条"对应的库**。

        例：记录 2013→A、2015→B、2018→C
            查 2016 → 落在 2015 与 2018 之间 → 本年库 B，前一年库**也是 B**
            查 2020 → 大于最大的 2018 → 本年库 C，前一年库**也是 C**
            查 2012 → 小于最小的 2013 → 无归属库（见下）

        【为什么前一年也同库？】
            这条规则表达的是"该库覆盖了从 iBeginYear 到下一条记录之前
            的全部年度"。所以 2016 与 2015 的数据物理上都在 B 库里，
            取前一年（2015）数据自然也从 B 库取。

        【查询年度早于最小 iBeginYear 怎么办？】
            说明该年度在账套里根本不存在（账套是 2013 年才建的，
            却要查 2012 年）。此时抛出错误，不猜库。

【只读承诺】
    本模块只执行 SELECT，不修改 UFSystem，也不修改任何业务库。
    （执行前还有 core/db.py 那三道只读防线兜底。）
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional

from core.db import Database, DatabaseError
from core.logging_setup import get_logger
from queries import u8

logger = get_logger(__name__)


class AccountResolveError(Exception):
    """账套/年度 解析失败。

    【新手必学：为什么要自定义异常类型？】
        调用方需要区分"解析不到账套库"和"其它数据库错误"：
        前者应该提示"请检查账套号与年度"，后者才提示"联系维护人员"。
        用一个专门的异常类型，就能精确地分别处理（见 app/generator.py）。
    """


@dataclass
class AccountEntry:
    """UA_AccountDatabase 中的一条记录（已规范化）。"""

    database: str            # 业务库名，如 "UFDATA_001" / "UFDATA_001_2026"
    begin_year: int          # 起始年度（判定归属的核心字段）
    end_year: Optional[int] = None   # 结束年度；本判定逻辑不依赖它，仅备查


@dataclass
class AccountTarget:
    """解析结果：本年度库 + 前一年度库，及其来源信息。

    【新手须知：为什么把 cDatabase 之外的东西也带回来？】
        只返回一个库名字符串当然也能用，
        但把"账套号、年度、覆盖区间、前一年库"一起带回来，就能写进
        日志与审计，日后核对"这个月到底读的是哪个库、为什么是它"时有据可查。
    """

    account_set: str        # 账套号，如 "001"
    year: int               # 会计年度，如 2026
    database: str           # 本年度业务库
    begin_year: Optional[int] = None      # 本年度库对应的起始年度
    end_year: Optional[int] = None        # 该记录的结束年度（备查）
    # ---- 前一年度 ----
    prev_year: Optional[int] = None       # 前一年度（= year - 1，无数据时为 None）
    prev_database: Optional[str] = None   # 前一年度业务库；None 表示"无前一年数据"
    prev_exact: bool = False              # 前一年库是否来自"精确年度匹配"
    # ---- 全部登记记录（供诊断与界面展示）----
    entries: list[AccountEntry] = field(default_factory=list)

    @property
    def has_prev(self) -> bool:
        """是否存在可用的前一年度数据。"""
        return bool(self.prev_database)

    @property
    def is_single_database(self) -> bool:
        """是否"单库"形态（该账套在登记表里只有一条记录）。

        【用途】提示用户"本账套为跨年度单库，各年度数据同库"。
        """
        return len(self.entries) == 1

    @property
    def is_cross_year(self) -> bool:
        """兼容旧名：等价于 is_single_database。

        【20261105 说明】
            早期版本用 iEndYear 是否为 NULL 判断"跨年度单库"，
            该理解有误（判定依据应为**记录条数**与 iBeginYear 关系）。
            为不破坏既有调用方，保留此属性名，语义改为"单库形态"。
        """
        return self.is_single_database

    def describe(self) -> str:
        """一句话描述，用于日志。

        【措辞讲究】用「本年度」「前一年度」两个明确标签，
            避免维护者在长日志里把两个库看混。
        """
        cur_span = (f"{self.begin_year} 年起" if self.begin_year is not None
                    else "区间未知")
        if self.prev_database:
            if self.prev_database == self.database:
                prev = f"前一年度 {self.prev_year} 年 → 同库"
            else:
                prev = f"前一年度 {self.prev_year} 年 → {self.prev_database}"
        else:
            prev = f"前一年度 {self.prev_year} 年 → 无（首年，无上年数据）"

        return (f"账套 {self.account_set} / {self.year} 年 "
                f"→ 业务库 {self.database}（{cur_span}，共 {len(self.entries)} 条登记）"
                f"；{prev}")


def resolve_account_database(db: Database, account_set: str, year: int,
                             system_database: str = "UFSystem",
                             switch: bool = True) -> AccountTarget:
    """解析（并可选切换）本年度业务库与前一年度业务库。

    参数：
        db              : 已连接到服务器的 Database 对象（当前在系统库上）。
        account_set     : 操作人选中的账套号，如 "001"。
        year            : 操作人选中的会计年度，如 2026。
        system_database : 系统库名，默认 UFSystem。
        switch          : 是否顺便把连接切到本年度业务库（默认切）。

    抛出 AccountResolveError：
        · 账套在 UA_AccountDatabase 中没有任何记录；
        · 查询年度早于该账套最早记录的起始年度；
        · 解析出的库名在服务器上并不存在。

    【新手难点：账套号为什么要规范化？】
        客户在界面里可能填 "1"、"01"、"001"、" 001 "。
        用友的 cAcc_Id 通常是定额宽度的字符串（如 3 位 "001"）。
        如果我们拿 "1" 直接去比，就会查不到 —— 客户明明填对了，却报"查无此账套"。
        所以这里统一：去空格 + 纯数字时左补零到 3 位。
    """
    acc = _normalize_account_set(account_set)
    year = int(year)

    logger.info("开始解析业务库：账套号=%s，年度=%s，系统库=%s",
                acc, year, system_database)

    # ---- 第1步：取回该账套的全部登记记录（不做年度过滤）----
    sql, params = u8.account_databases(system_database, acc)
    try:
        rows = db.query(sql, params, tag="ACCOUNT")
    except DatabaseError as exc:
        # 【重要：把技术错误翻译成人话，且原始错误不往界面传】
        #   最常见的两种现场情况：
        #     1) 连接的不是系统库 → 找不到 UA_AccountDatabase；
        #     2) 用的是普通账套操作员账号，没有系统库的查询权限。
        logger.error("读取账套注册信息失败：%s", exc, exc_info=True)
        raise AccountResolveError(
            "无法读取用友的账套注册信息。\n"
            "常见原因：登录账号没有系统库的查询权限，或系统库名不是 UFSystem。\n"
            "请联系系统管理员确认数据库账号权限后重试。"
        ) from exc

    if not rows:
        # 查不到就是查不到，绝不"猜一个默认库"。
        hint = _account_hint(db, system_database, acc)
        raise AccountResolveError(
            f"未找到账套 {acc} 的任何数据库登记记录。\n"
            "请核对：①账套号是否正确；②该账套是否已在本服务器上建立。"
            f"{hint}"
        )

    entries = _parse_entries(rows)
    if not entries:
        raise AccountResolveError(
            f"账套 {acc} 有登记记录，但库名或起始年度无法识别。\n"
            "这属于账套登记信息异常，请让系统管理员检查用友系统库中的账套登记。"
        )

    # ---- 第2步：按现场规则推导"本年度库"与"前一年度库" ----
    current, prev, prev_exact = _pick_databases(entries, year, acc)

    if current is None:
        earliest = entries[0].begin_year
        raise AccountResolveError(
            f"账套 {acc} 的最早起始年度为 {earliest} 年，"
            f"但您选择的年度是 {year} 年。\n"
            f"该年度早于账套建立时间，账套中不存在这一年的数据。\n"
            f"目前该账套可查询的年度范围：{earliest} 年及以后。"
        )

    target = AccountTarget(
        account_set=acc, year=year,
        database=current.database,
        begin_year=current.begin_year, end_year=current.end_year,
        prev_year=year - 1 if prev is not None else None,
        prev_database=prev.database if prev is not None else None,
        prev_exact=prev_exact,
        entries=entries,
    )
    logger.info("业务库解析成功：%s", target.describe())

    # ---- 第3步：验证库真的存在（含前一年度库）----
    _ensure_exists(db, current.database, acc, year, "本年度")
    if prev is not None and prev.database != current.database:
        _ensure_exists(db, prev.database, acc, year - 1, "前一年度")

    if switch:
        db.switch_database(current.database)
        logger.info("取数连接已固定到业务库：%s（后续所有查询默认在此库）",
                    current.database)

    return target


def _parse_entries(rows) -> list[AccountEntry]:
    """把 query 返回的原始行整理成 AccountEntry 列表，并按起始年度升序。

    【为什么要过滤无效行？】
        UA_AccountDatabase 里可能存在库名为空、起始年度为 NULL 的脏记录。
        这类记录无法参与判定，留着只会干扰排序与相邻关系，故直接跳过并告警。
    """
    entries: list[AccountEntry] = []
    for r in rows:
        try:
            database = str(r[0]).strip() if r[0] else ""
            begin_raw = r[1]
            end_raw = r[2] if len(r) > 2 else None
            if not database or begin_raw is None:
                logger.warning("跳过无效账套登记记录（库名或起始年度为空）：%r", r)
                continue
            entries.append(AccountEntry(
                database=database,
                begin_year=int(begin_raw),
                end_year=int(end_raw) if end_raw is not None else None,
            ))
        except (IndexError, TypeError, ValueError) as exc:
            logger.warning("跳过无法解析的账套登记记录 %r：%s", r, exc)

    entries.sort(key=lambda e: e.begin_year)
    return entries


def _pick_databases(entries: list[AccountEntry], year: int, acc: str
                    ) -> tuple[Optional[AccountEntry], Optional[AccountEntry], bool]:
    """按现场规则选出本年度记录与前一年度记录。

    返回 (current, prev, prev_exact)：
        current     : 本年度对应的记录；None 表示查询年度早于账套起始年
        prev        : 前一年度对应的记录；None 表示"无前一年数据"
        prev_exact  : 前一年库是否来自"精确年度匹配"（用于日志区分）

    【规则实现说明（对应模块头的三种情形）】
        情形一 N==1        : 本年、前一年都是这唯一一条
        情形二 年度==某条  : 本年即该条；前一年取相邻较小的一条（若无则 None）
        情形三 年度落在区间: 本年取"较小者"那条；前一年也取同一条
    """
    if not entries:
        return None, None, False

    # ---- 情形一：只有一条记录（跨年度单库 / 仅一年）----
    if len(entries) == 1:
        only = entries[0]
        # 同期判断：查询年度不能早于该库起始年度，否则是"账套还没建"
        if year < only.begin_year:
            return None, None, False
        return only, only, True

    # ---- 情形二：查询年度恰好等于某条记录的起始年度 ----
    exact = [e for e in entries if e.begin_year == year]
    if exact:
        current = exact[0]
        idx = entries.index(current)
        if idx == 0:
            # 已是首年（最小起始年度）→ 没有前一年数据
            return current, None, False
        # 前一年取"相邻但较小"的那条
        return current, entries[idx - 1], True

    # ---- 情形三：查询年度未精确命中任何起始年度 ----
    # 找出所有起始年度 <= year 的记录，取最大的那条（即"较小的相邻者"）
    #
    # 【这一行同时覆盖两种子情形，务必分清】
    #   (a) 查询年度落在两条记录之间，如记录 2015→B、2018→C，查 2016
    #       → lower = [2015, 2018 之前那些]，末位 2015 → 用 B
    #   (b) 查询年度**大于所有记录**，如记录止于 2018，查 2020
    #       → lower 含全部记录，末位 2018 → 用 C
    #
    # 【(b) 为什么也取"最大那条"，而不是报错？—— 20261105 现场确认】
    #   理由：该库可能一直延用（账套停用后登记未清理，或升级后沿用旧库）。
    #   若此处报错，会挡住"账套确实还在用、只是没补登记"的正常场景。
    #   代价是：若账套真的在 2018 年后停用，用户查 2020 会连上 C 库、
    #   取到空数据、报表填出一堆 0 —— 因此**年度是否有数据由用户自行核对**。
    #
    #   【不要"顺手改成报错"，也不要改成取最后一条之外的记录】
    #   这个取舍已经现场确认过；测试 test_after_last_record 锁定该行为。
    lower = [e for e in entries if e.begin_year <= year]
    if not lower:
        # 查询年度早于所有记录 → 账套尚未建立（与 (b) 情形相反，必须报错）
        return None, None, False

    current = lower[-1]   # entries 已升序，lower 亦升序，末位即最大
    # 【关键】同日历年度的前一年也在这个库里（该库覆盖到下一记录之前）
    return current, current, False


def _ensure_exists(db: Database, database: str, acc: str, year: int,
                   label: str) -> None:
    """确认库在服务器上真实存在，否则抛出可读错误。

    【为什么必须先查存在性，不能直接 USE？】
        UA_AccountDatabase 里可能有"历史残留"记录：账套早就删了/迁走了，
        但注册表里还留着库名。直接 USE 会抛一句英文报错，客户看不懂。
        提前查一次，提示就能精准到"哪个库、哪一年、请联系管理员"。
    """
    if db.database_exists(database):
        return
    raise AccountResolveError(
        f"账套 {acc} 的{label}（{year} 年）对应业务库 {database}，\n"
        "但服务器上并不存在这个库（可能已删除或已升级迁移）。\n"
        "请联系系统管理员核实账套登记信息。"
    )


def _normalize_account_set(value: str) -> str:
    """把用户输入的账套号规范化成用友习惯的写法。

    【转换示例】
        "1"      → "001"
        "01"     → "001"
        " 001 "  → "001"
        "0001"   → "0001"（本来就超过 3 位，不动它）
        ""       → 抛错（空账套号无意义）
    """
    text = str(value or "").strip()
    if not text:
        raise AccountResolveError("账套号为空，请先填写账套号（如 001）。")
    if text.isdigit() and len(text) < 3:
        return text.zfill(3)
    return text


def _account_hint(db: Database, system_database: str, acc: str) -> str:
    """查不到账套时，顺手列出现有账套，帮人一眼看出"是不是填错了"。

    【注意】这里所有异常都不往外抛：
        一个"辅助提示"失败，绝不能把主流程带崩。
    """
    try:
        sql, params = u8.list_accounts(system_database)
        rows = db.query(sql, params, tag="ACCOUNT")
    except Exception:  # noqa: BLE001
        return ""

    if not rows:
        return "\n（系统库中没有任何账套登记，请确认连接的是用友系统库。）"

    ids: list[str] = []
    for r in rows:
        try:
            v = str(r[0]).strip()
        except (IndexError, TypeError):
            continue
        if v and v not in ids:
            ids.append(v)
    shown = "、".join(ids[:12])
    more = "…" if len(ids) > 12 else ""
    return f"\n系统中现有账套号：{shown}{more}"