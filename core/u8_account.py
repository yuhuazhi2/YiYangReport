# -*- coding: utf-8 -*-
"""
账套 → 年度库 解析器（本项目最关键的一步）
============================================================

【模块职责】
    给定"账套号 + 会计年度"，从用友系统库 UFSystem 的
    UA_AccountDatabase 表里查出**应该连哪个年度库**（如 UFDATA_001_2026）。

【为什么这一步必须单独做一个模块？】
    因为它是"数据对不对"的总开关。
    只要这一步错了，后面 43 张报表算得再准也是错的账 —— 而且
    程序**不会报错**，会安安静静地输出另一家公司的数字。
    这种错误比崩溃危险一百倍：崩溃你立刻知道，错账你可能一个月后才发现。

    所以它被单独拆出来，做到三件事：
        ① 规则只有一份（就在这里），不散落在各处；
        ② 出错时给出**人能看懂、能照着处理**的提示，而不是 SQL 报错；
        ③ 结果在切换数据库前还要再验证一次"库真的存在"。

【整体流程（你在日志里能看到对应的行）】
    登录服务器（连系统库 UFSystem）
        ↓
    查 UA_AccountDatabase：账套号 + 年度 → cDatabase
        ↓
    验证该库在服务器上真实存在
        ↓
    把连接切到该库
        ↓
    全部 43 张报表都从这个库取数

【只读承诺】
    本模块只执行 SELECT，不修改 UFSystem，也不修改任何年度库。
    （执行前还有 core/db.py 那三道只读防线兜底。）
"""
from __future__ import annotations

from dataclasses import dataclass
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
class AccountTarget:
    """解析结果：最终要连的年度库及其来源信息。

    【新手须知：为什么把 cDatabase 之外的东西也带回来？】
        只返回一个库名字符串当然也能用，
        但把"账套号、年度、覆盖区间"一起带回来，就能写进日志与审计，
        日后核对"这个月到底读的是哪个库、为什么是它"时有据可查。
    """

    account_set: str        # 账套号，如 "001"
    year: int               # 会计年度，如 2026
    database: str           # 年度库名，如 "UFDATA_001_2026"
    begin_year: Optional[int] = None   # 该库覆盖的开始年度（来自 iBeginYear）
    end_year: Optional[int] = None     # 该库覆盖的结束年度（来自 iEndYear，可能为空）

    def describe(self) -> str:
        """一句话描述，用于日志。

        【措辞讲究】外层模板里已经有一对括号「（覆盖 …）」，
            所以这里的区间描述**不能再套括号**，否则读起来是
            「（覆盖 2026（仅该年度））」—— 两层括号夹在一起，维护者一眼扫过去会卡住。
            写成「（覆盖 2026 年，仅该年度）」就通顺了。
        """
        if self.end_year is None:
            # iEndYear 为空的含义：这条登记只代表开始那一年（不是"从此一直有效"）
            span = f"{self.begin_year} 年，仅该年度"
        elif self.end_year != self.begin_year:
            span = f"{self.begin_year}~{self.end_year} 年"
        else:
            span = f"{self.begin_year} 年"
        return (f"账套 {self.account_set} / {self.year} 年 "
                f"→ 年度库 {self.database}（覆盖 {span}）")


def resolve_account_database(db: Database, account_set: str, year: int,
                             system_database: str = "UFSystem",
                             switch: bool = True) -> AccountTarget:
    """解析并（可选）切换连接，返回最终要用的年度库。

    参数：
        db              : 已连接到服务器的 Database 对象（当前在系统库上）。
        account_set     : 操作人选中的账套号，如 "001"。
        year            : 操作人选中的会计年度，如 2026。
        system_database : 系统库名，默认 UFSystem。
        switch          : 是否顺便把连接切到该年度库（默认切）。

    抛出 AccountResolveError：
        账套号+年度在 UA_AccountDatabase 里找不到匹配记录，
        或者查到了库名但服务器上并没有这个库。
        这两种都属于"需要人去核对"的问题，所以单独抛出。

    【新手难点：账套号为什么要规范化？】
        客户在界面里可能填 "1"、"01"、"001"、" 001 "。
        用友的 cAcc_Id 通常是定额宽度的字符串（如 3 位 "001"）。
        如果我们拿 "1" 直接去比，就会查不到 —— 客户明明填对了，却报"查无此账套"。
        所以这里统一：去空格 + 纯数字时左补零到 3 位。
        （只补到 3 位，不改变本来就超过 3 位的账套号。）
    """
    acc = _normalize_account_set(account_set)

    logger.info("开始解析账套库：账套号=%s，年度=%s，系统库=%s",
                acc, year, system_database)

    sql, params = u8.account_database(system_database, acc, int(year))
    try:
        rows = db.query(sql, params, tag="ACCOUNT")
    except DatabaseError as exc:
        # 【重要：把技术错误翻译成人话，且原始错误不往界面传】
        #   最常见的两种现场情况：
        #     1) 连接的不是系统库 → 找不到 UA_AccountDatabase；
        #     2) 用的是普通账套操作员账号，没有系统库的查询权限。
        #   这两种都不该让客户看到英文堆栈，也不该把库名/表名/权限对象名
        #   抛到界面上（那是内部结构）。所以：完整异常只写 logs/app.log。
        logger.error("读取账套注册信息失败：%s", exc, exc_info=True)
        raise AccountResolveError(
            "无法读取用友的账套注册信息。\n"
            "常见原因：登录账号没有系统库的查询权限，或系统库名不是 UFSystem。\n"
            "请联系系统管理员确认数据库账号权限后重试。"
        ) from exc

    if not rows:
        # 查不到就是查不到，绝不"猜一个默认库"。
        # 【设计原则】宁可停下来让人核对，也不能默默查错账套。
        hint = _account_hint(db, system_database, acc)
        raise AccountResolveError(
            f"未找到匹配的账套库：账套号 {acc}、年度 {year}。\n"
            "请核对：①账套号是否正确；②该年度是否已建立账套库。"
            f"{hint}"
        )

    # 理论上"账套+年度"应唯一；万一有多条（历史遗留），取覆盖区间最新的那条
    if len(rows) > 1:
        logger.warning("账套 %s 在 %s 年匹配到 %d 条记录，取第一条（区间最新）：%s",
                       acc, year, len(rows), rows)

    first = rows[0]
    # 只取第一列（cDatabase）。这里用取下标而不是 __getitem__ 的字段名，
    # 因为 query() 返回的是普通元组，不是字典。
    database = str(first[0]).strip() if first and first[0] else ""
    if not database:
        raise AccountResolveError(
            f"账套 {acc} 在 {year} 年的记录存在，但对应的年度库名为空。\n"
            "这属于账套登记信息异常，请让系统管理员检查用友系统库中的账套登记。"
        )

    begin_year = end_year = None
    try:
        begin_year = int(first[1]) if first[1] is not None else None
        end_year = int(first[2]) if first[2] is not None else None
    except (IndexError, TypeError, ValueError):
        # 区间信息只是"用于展示与日志"，取不到不影响主流程
        pass

    target = AccountTarget(
        account_set=acc, year=int(year), database=database,
        begin_year=begin_year, end_year=end_year,
    )
    logger.info("账套库解析成功：%s", target.describe())

    # ---- 第二道保险：库名解析出来了，还要确认它真的存在 ----
    # 【为什么还要查一次？】
    #   UA_AccountDatabase 里可能有"历史残留"记录：账套早就删了，
    #   但注册表里还留着 UFDATA_001_2019 这样的名字。
    #   如果直接 USE，SQL Server 会报"无法打开数据库"，
    #   客户看到的就是一句莫名其妙的英文。提前查一下，提示就精准了。
    if not db.database_exists(database):
        # 【提示文案里的库名要不要给客户看？】
        #   这里保留库名，是因为这句话的用途是"让系统管理员去核实"，
        #   而管理员核实时必须有库名才查得下去（少了它等于没说）。
        #   同时这句话会原样显示在客户区，客户看不懂也无妨 ——
        #   它明确写着"请联系系统管理员"，把动作指向了正确的人。
        raise AccountResolveError(
            f"账套 {acc} 在 {year} 年对应的年度库 {database}，\n"
            "但服务器上并不存在这个库（可能已删除或已升级迁移）。\n"
            "请联系系统管理员核实账套登记信息。"
        )

    if switch:
        db.switch_database(database)
        logger.info("取数连接已固定到年度库：%s（后续所有查询都在此库）", database)

    return target


def _normalize_account_set(value: str) -> str:
    """把用户输入的账套号规范化成用友习惯的写法。

    【转换示例】
        "1"      → "001"
        "01"     → "001"
        " 001 "  → "001"
        "0001"   → "0001"（本来就超过 3 位，不动它）
        ""       → 抛错（空账套号无意义）

    【新手必学：str.zfill(3)】
        它把字符串左边补 0 到至少 3 位："1".zfill(3) → "001"。
        但它只对"纯数字"有意义，所以要先判断 isdigit()。
        如果账套号里含字母（极少数定制环境），就原样保留、只去空格。
    """
    text = str(value or "").strip()
    if not text:
        raise AccountResolveError("账套号为空，请先填写账套号（如 001）。")
    if text.isdigit() and len(text) < 3:
        return text.zfill(3)
    return text


def _account_hint(db: Database, system_database: str, acc: str) -> str:
    """查不到账套时，顺手列出现有账套，帮人一眼看出"是不是填错了"。

    【设计意图：把"报错"变成"帮他改对"】
        客户看到"未找到账套库"时，第一反应是"我填的没错啊"。
        如果同时把系统里实际存在的账套号列出来，他一眼就能发现
        "哦，我要的是 002，不是 001"。
        这比让他去找维护人员快得多，也减少我们的支持成本。

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

    # 去重后取前若干个账套号，避免提示过长把界面撑爆
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