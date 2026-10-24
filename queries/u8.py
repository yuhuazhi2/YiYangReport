# -*- coding: utf-8 -*-
"""
【模块职责】用友 U8 V18.0 后台取数 SQL 定义

本模块把"要用到的 SQL 语句"集中放在一个文件里。

【新手必学】为什么要单独搞一个 queries/ 目录、不把 SQL 散落在各处？
    1. 好找：所有 SQL 都在这一处，想改某个表的取数逻辑，直接来这里找
    2. 好审：财务/IT 人员核对"程序到底查了哪些表、哪些字段"时，
       只要读这一个文件就够了，不用翻遍整个项目
    3. 好换：将来 U8 升级、字段改名，改动只集中在这里

【本模块的核心设计：每个函数返回 (sql, params) 元组】
    这是给初级开发者重点讲解的一个设计模式，好处在于：

        ① 写 SQL 和跑 SQL 彻底分开
           这里只负责"拼出 SQL 语句 + 准备参数"，不负责连接数据库。
           真正执行由 core/db.py 的 Database.query() 来做。

        ② 方便复用与测试
           想看看某个函数生成的 SQL 长什么样？直接 print 就行，
           不需要真的连上数据库。这叫"纯函数"——给定输入必得同样输出。

        ③ 防止 SQL 注入
           注意下面所有 SQL 都写 "WHERE iyear = ?"，而不是直接把 year
           拼进字符串。? 是"占位符"，真正的值放在 params 列表里，
           由数据库驱动去安全替换。这样即使用户输入了恶意的字符串，
           也只会被当成"一个值"，而不会被当成"SQL 代码"执行。

【新手必学】推荐使用方式（伪代码）：
        sql, params = gl_account_summary(2026, 3)
        rows = db.query(sql, params)     # 交给 Database 去执行

⚠⚠⚠ 极其重要的现场提醒 ⚠⚠⚠
    本文件中的表名/字段名是**用友 U8 的通用命名**，仅用于搭框架和演示。
    实际部署到客户现场时，**必须**打开该客户真实的 UFDATA 年度库，
    用 core/db.py 的 list_columns() 核对每一张表的真实字段名再调整！
    不同版本、不同模块启用情况的 U8，字段可能有出入。
    千万不要以为这份 SQL 拿去就能跑出正确的数。

【新手须知】组织方式约定：
    - 每个函数对应"一类数据需求"，返回 (sql, params) 元组
    - 期间参数统一用 (iyear, iperiod)，即"会计年、会计期间(月)"
    - 注意区分两种库：
        数据字典库 UFSystem    —— 存账套信息、用户、权限等系统级数据
        年度库     UFDATA_xxx_yyyy —— 每个账套每年一个库，业务数据都在这里
"""
from __future__ import annotations

# 【新手必学】typing.Any 表示"任意类型"。
# 这里用它标注 params 列表（因为里面可能装 int、str 等混合类型）。
from typing import Any


# ---------------------------------------------------------------------------
# 一、总账类
# ---------------------------------------------------------------------------
def gl_account_summary(year: int, month: int) -> tuple[str, list]:
    """
    科目汇总表（GL_accsum）：取某期间各科目借贷方发生额与期末余额。
    用于资产负债表、利润表等。

    【新手必学】函数签名 -> tuple[str, list] 是"类型提示"：
        明确告诉调用者"我会返回一个元组，第一个元素是字符串(SQL)，
        第二个元素是列表(参数)"。Python 不会强制检查，但它提高了
        代码可读性，写代码时编辑器还能给出智能提示。

    【字段小抄】GL_accsum 的字段含义：
        ccode    科目编码
        cbegind_c 期初余额（方向 + 金额，'c' 代表方向）
        mb       期初余额    md  借方发生额
        mc       贷方发生额  me  期末余额
    """
    # 【新手必学】三个双引号包起来的是"多行字符串"。
    # 写 SQL 时特别有用，可以随便换行、缩进，不用拼一堆 + 号。
    sql = """
        SELECT ccode, cbegind_c, mb, md, mc, me
        FROM GL_accsum
        WHERE iyear = ? AND iperiod = ?
    """
    # 【新手必学】return 时用逗号隔开，就是返回一个元组（tuple）。
    # 注意：参数顺序必须和 SQL 里 ? 出现的顺序**一一对应**！
    # 这里第一个 ? 对应 year，第二个 ? 对应 month，千万别搞反。
    return sql, [year, month]


def gl_vouchers(year: int, month: int) -> tuple[str, list]:
    """凭证分录（GL_accvouch）：取指定期间已记账凭证"""
    sql = """
        SELECT ccode, cdigest, md, mc, cdept_id, cperson_id, ccus_id,
               csup_id, citem_id, cbook, ibook, csign, isignseq
        FROM GL_accvouch
        WHERE iyear = ? AND iperiod = ? AND cbook IS NOT NULL
    """
    # 【新手须知】"cbook IS NOT NULL" 是 SQL 里判断"不为空"的写法。
    # 这里用来筛掉"未记账凭证"（U8 里 cbook 为空表示尚未记账）。
    # 【新手难点】SQL 里判断空值必须用 "IS NULL" / "IS NOT NULL"，
    # 不能写 "= NULL"（那样永远不成立），这是新手极容易踩的坑。
    return sql, [year, month]


def gl_assist_summary(year: int, month: int) -> tuple[str, list]:
    """辅助核算汇总（GL_accass）

    【新手须知】"辅助核算"是用友的概念：除了会计科目，还按
    部门/个人/客户/供应商/项目等维度细分。这张表就是各维度的汇总。
    """
    sql = """
        SELECT ccode, cdept_id, cperson_id, ccus_id, csup_id,
               mb, md, mc, me
        FROM GL_accass
        WHERE iyear = ? AND iperiod = ?
    """
    return sql, [year, month]


def gl_subject_codes() -> tuple[str, list]:
    """科目档案（Code）

    【新手难点】注意本函数**没有参数**（括号是空的）！
    所以它返回的 params 是个**空列表** []，表示"这条 SQL 不需要占位符"。
    这是一种统一约定：无论有没有参数，都返回元组，
    调用方就不用写两套代码去判断了。
    """
    sql = """
        SELECT ccode, ccode_name, cclass, bperson, bdept, bcus, bsup,
               bitem, bend
        FROM Code
        WHERE bend = 0
        ORDER BY ccode
    """
    # 【新手必学】bend = 0 表示"未封存/在用"的科目（U8 惯例：0=使用中）。
    # ORDER BY ccode 让结果按科目编码升序排列，方便人工查看。
    return sql, []


# ---------------------------------------------------------------------------
# 二、货币资金 / 票据
# ---------------------------------------------------------------------------
def bank_deposits(year: int, month: int) -> tuple[str, list]:
    """银行存款日记账（按科目 1002 取数）

    【新手必学】LIKE '1002%' 是"模糊匹配"：
        % 是通配符，代表"任意多个字符"。
        所以 '1002%' 能匹配 1002、100201、10020101 等所有以 1002 开头的科目。
        （1002 是企业会计准则里的"银行存款"一级科目编码）

    【新手难点】为什么这里也要用 ? 占位符，LIKE 的值却直接写死？
        因为 '1002%' 是**程序写死的常量**，不是用户输入，不存在注入风险；
        而 year/month 来自用户输入，所以必须走占位符。
    """
    sql = """
        SELECT ccode, cdigest, md, mc, cbook
        FROM GL_accvouch
        WHERE iyear = ? AND iperiod = ?
          AND ccode LIKE '1002%'
    """
    return sql, [year, month]


# ---------------------------------------------------------------------------
# 三、往来类（应收/应付）
# ---------------------------------------------------------------------------
def ar_aging(year: int, month: int) -> tuple[str, list]:
    """应收账款明细 + 辅助（通过 GL_accass 取客户维度）

    【新手必学】LEFT JOIN 是"左连接"：
        以左边的表（GL_accass，记为 a）为基准，把右边
        （Customer，记为 c）里匹配得上的行拼进来；
        就算右表没匹配上，左表的数据也不会丢（右表那几列显示为空）。

    【新手必学】"表别名"（a、c）的作用：
        给表起个短名字，后面字段就可以写成 a.ccode、c.ccus_name，
        既简洁，又能避免两张表有同名字段时搞混。
    """
    sql = """
        SELECT a.ccode, a.ccus_id, c.ccus_name,
               a.md, a.mc, a.me
        FROM GL_accass a
        LEFT JOIN Customer c ON a.ccus_id = c.ccus_id
        WHERE a.iyear = ? AND a.iperiod = ?
          AND a.ccode LIKE '1122%'
    """
    # 【新手须知】1122 = "应收账款"一级科目编码。
    return sql, [year, month]


def ap_aging(year: int, month: int) -> tuple[str, list]:
    """应付账款明细 + 辅助（供应商维度）

    【新手必学】本函数与上面的 ar_aging 结构几乎一样，只换了
        两张表（Vendor）和科目编码（2202 = 应付账款）。
        如果你要新增一类往来核算，照抄这个函数改改即可 ——
        这就是把 SQL 集中管理带来的好处。
    """
    sql = """
        SELECT a.ccode, a.csup_id, s.csup_name,
               a.md, a.mc, a.me
        FROM GL_accass a
        LEFT JOIN Vendor s ON a.csup_id = s.csup_id
        WHERE a.iyear = ? AND a.iperiod = ?
          AND a.ccode LIKE '2202%'
    """
    return sql, [year, month]


# ---------------------------------------------------------------------------
# 四、存货类
# ---------------------------------------------------------------------------
def inventory_balance(year: int, month: int) -> tuple[str, list]:
    """存货现存量 / 结存（ST_CurrentStock 或汇总表）

    【新手必学】cInvCode = 存货编码，cWhCode = 仓库编码。
        U8 里字段名前缀有讲究：
            c 开头 → 字符型(char)
            i 开头 → 整型(int)
            d 开头 → 日期型(date)
            b 开头 → 布尔型(bit，0/1)
        看懂前缀，猜字段含义会快很多。
    """
    sql = """
        SELECT cInvCode, iQuantity, iNum, cWhCode
        FROM ST_CurrentStock
        WHERE iYear = ? AND iMonth = ?
    """
    return sql, [year, month]


def inventory_archive() -> tuple[str, list]:
    """存货档案（Inventory）

    【新手必学】这是"档案表"（基础资料），不带期间参数，
        因为档案是长期有效的，不随月份变化（和汇总表不同）。
    """
    sql = """
        SELECT cInvCode, cInvName, cInvStd, cInvCCode, cComUnitCode
        FROM Inventory
        WHERE bInvType = 1
    """
    # 【新手必学】bInvType = 1 表示"存货"类型（区别于 2=劳务 等）。
    return sql, []


def rd_records(year: int, month: int) -> tuple[str, list]:
    """收发记录主表（RdRecord），用于出入库汇总

    【新手难点】这里用的是 YEAR(dDate) 和 MONTH(dDate) 函数，
        即"从日期字段里抽出年份和月份"，而不是直接等于某个期间字段。
        为什么？因为 RdRecord 只存日期（dDate），没有单独的"年月"列，
        所以只能把日期拆开来比。

    【新手警告】⚠ 这种写法有个性能陷阱：
        对字段用函数（YEAR(dDate)）会导致**索引失效**，表一大就会变慢。
        更快的写法通常是范围比较，例如：
            WHERE dDate >= ? AND dDate < ?
        （传两个日期边界值进 params）
        本框架为演示清晰先保留直观写法，实际部署时建议按上述方式优化。
    """
    sql = """
        SELECT ID, cCode, cWhCode, dDate, cVouchType
        FROM RdRecord
        WHERE YEAR(dDate) = ? AND MONTH(dDate) = ?
    """
    return sql, [year, month]


# ---------------------------------------------------------------------------
# 五、固定资产
# ---------------------------------------------------------------------------
def fa_cards(year: int, month: int) -> tuple[str, list]:
    """固定资产卡片（fa_Cards）

    【新手难点】注意这个函数的参数用法和别的不同！
        其他函数都是 [year, month] 两个参数，这里却**只用了一个 month**：

            WHERE MONTH(dStartDate) <= ?
            return sql, [month]

        这是因为固定资产卡片的逻辑是"开始使用月份 <= 当前月份"，
        也就是说"到这个月为止已经存在的资产都算进来"。
        这是业务逻辑决定的，不是笔误。改代码前一定要想清楚业务含义。

    【新手警告】⚠ 参数签名里保留了 year，但 SQL 里没用到。
        这属于"为了与其他函数签名统一"而保留的形参。
        若你确认不需要，可以删掉；但删了要同步改所有调用处，谨慎操作。
    """
    sql = """
        SELECT sCardNum, sAssetNum, sDeptNum, dStartDate, dblValue,
               dblDep, dblThisDep, dblMonth, sClassNum
        FROM fa_Cards
        WHERE MONTH(dStartDate) <= ?
    """
    return sql, [month]


def fa_depreciation(year: int, month: int) -> tuple[str, list]:
    """固定资产折旧（fa_Depreciation）"""
    sql = """
        SELECT sCardNum, iPeriod, dblDep, dblThisDep
        FROM fa_Depreciation
        WHERE iYear = ? AND iPeriod = ?
    """
    return sql, [year, month]


# ---------------------------------------------------------------------------
# 六、薪资类
# ---------------------------------------------------------------------------
def salary_data(year: int, month: int, wa_class: str = "") -> tuple[str, list]:
    """工资数据（wa_GZData）。wa_class 为工资类别编码

    【新手必学】wa_class: str = "" 是"默认参数"：
        调用时可以不传这个参数（如 salary_data(2026, 3)），
        此时 wa_class 自动等于 ""（空字符串）。

    【新手必学】本函数演示了"根据条件动态拼 SQL"的经典写法：
        - 先拼好基础部分
        - 满足条件时，再追加一段 AND ...
        - 同时往 params 里追加对应的参数
        注意：SQL 里新增了几个 ?，params 就必须跟着加几个值，顺序要对齐。
    """
    sql = """
        SELECT cgzgrade, cgzitem, cDeptNum, cPsnNum, fValue
        FROM wa_GZData
        WHERE iYear = ? AND iMonth = ?
    """
    # 【新手须知】params: list[Any] = [year, month] 是先声明类型再赋值，
    # 意思和 params = [year, month] 一样，只是多写了类型标注。
    params: list[Any] = [year, month]
    # 【新手必学】在 Python 里，空字符串 "" 在 if 中会被判为 False。
    # 所以 "if wa_class:" 的意思是"只有传了非空的工资类别时才追加条件"。
    if wa_class:
        # 【新手必学】sql += "..." 是"在原字符串后面追加"，等价于
        # sql = sql + "..."。字符串是不可变的，这实际上是造了个新字符串
        # 再赋回给 sql 变量。
        sql += " AND cGZGrade = ?"
        # 【新手必学】.append(x) 往列表末尾添加一个元素。
        # 这里必须紧随 SQL 里的 ? 一起加，保证顺序一致。
        params.append(wa_class)
    return sql, params


# ---------------------------------------------------------------------------
# 七、账套注册信息（系统库 UFSystem）—— 决定"该连哪个年度库"
# ---------------------------------------------------------------------------
def account_database(system_database: str, acc_id: str, year: int) -> tuple[str, list]:
    """查询"某账套在某年度"对应的年度库名（如 UFDATA_001_2026）。

    【这张表是什么？】
        用友把"有哪些账套、每个账套有哪几年的库"登记在系统库
        UFSystem 的 UA_AccountDatabase 表里。这是**权威来源**：
        年度库叫什么名字，必须问它，不能靠人猜、更不能靠客户手填。

    【关键字段】
        cAcc_Id     账套号（如 "001"），对应界面上的"账套号"
        cDatabase   年度库名（如 "UFDATA_001_2026"）—— 我们要的就是它
        iBeginYear  该年度库覆盖的开始年度
        iEndYear    该年度库覆盖的结束年度；**可能是 NULL**

    【匹配规则（现场确定的业务口径，不要随意改动）】
        条件一：cAcc_Id 等于操作人选的账套号
        条件二：满足下面任意一条
            a) 选的年度在 iBeginYear 与 iEndYear 之间（含两端）
            b) iEndYear 为 NULL，且选的年度正好等于 iBeginYear

    【新手难点：为什么 (b) 里只用 = 而不是 >=？】
        按现场给的口径：iEndYear 为 NULL 时，这条记录**只代表开始那一年**，
        而不是"从开始年度一直有效"。这是用友数据本身的约定，
        我们必须照着它判断，否则可能把 2026 年的账错当成 2030 年的。

    【新手必学：为什么库名要拼进 SQL，而不是用 ? 占位符？】
        因为 [UFSystem] 是"库名/表名"，属于标识符；
        占位符 ? 只能放"值"（如账套号、年度），不能放标识符。
        所以库名用 quote_ident 安全包裹后拼接，而**账套号与年度仍走 ?**，
        这样既满足语法要求，又杜绝了注入风险。

    【新手必学：(a) 与 (b) 用 OR 连接，外面套一层括号】
        WHERE cAcc_Id = ?
          AND ( (年度区间判断) OR (NULL 特例判断) )
        如果不加外层括号，SQL 的 AND 优先级高于 OR，
        就会变成 "A AND B OR C"，与我们的本意完全不同 —— 这是极常见的错误。
    """
    # 延迟导入：避免 queries 与 core 之间产生循环导入
    from core.db import quote_ident

    sysdb = quote_ident(system_database)
    # 【新手须知：为什么把 iBeginYear / iEndYear 也一起查出来？】
    #   它们不影响"取哪个库"的判断（那是 WHERE 的事），
    #   但带回上层后，可以写进日志："账套 001 的 2026 年 → UFDATA_001_2026
    #   （覆盖 2026~2026）"。日后核对"为什么查的是这个库"时，
    #   这行日志就是最直接的证据。
    sql = f"""
        SELECT cDatabase, iBeginYear, iEndYear
        FROM {sysdb}.dbo.UA_AccountDatabase
        WHERE cAcc_Id = ?
          AND (
                (iEndYear IS NOT NULL AND ? BETWEEN iBeginYear AND iEndYear)
             OR (iEndYear IS NULL AND ? = iBeginYear)
          )
        ORDER BY iBeginYear DESC
    """
    # 三个 ? 依次是：账套号、年度（区间判断）、年度（NULL 特例判断）
    return sql, [acc_id, year, year]


def list_accounts(system_database: str) -> tuple[str, list]:
    """列出所有账套（供界面辅助提示、以及维护时排查"账套号填错了"）。

    【用途】客户说"账套号是 001"但查不到库时，用它把所有账套号列出来，
        一看就知道是不是输成了 "1"、或者实际叫 "001"。
    【注意】只是"读"，不涉及任何写入。
    """
    from core.db import quote_ident

    sysdb = quote_ident(system_database)
    sql = f"""
        SELECT DISTINCT cAcc_Id, cDatabase, iBeginYear, iEndYear
        FROM {sysdb}.dbo.UA_AccountDatabase
        ORDER BY cAcc_Id, iBeginYear
    """
    return sql, []


# ---------------------------------------------------------------------------
# 八、元数据 / 诊断（供核对与排错）
# ---------------------------------------------------------------------------
def check_table_exists(table_name: str) -> tuple[str, list]:
    """检查表是否存在（用于取数前校验）

    【新手必学】OBJECT_ID(?) 是 SQL Server 的内置函数：
        传入表名，返回该表在数据库里的对象编号；
        表不存在时返回 NULL。常用来在取数前"探一探表在不在"，
        避免直接查一个不存在的表而报错。
    """
    sql = "SELECT OBJECT_ID(?) AS oid"
    return sql, [table_name]


def sample_rows(table_name: str, n: int = 10) -> tuple[str, list]:
    """取样若干行（用于人工核对后台数据）

    【新手难点】这是全文件唯一"直接把变量拼进 SQL"的地方，
        看起来似乎违反了前面说的"必须用 ? 占位符"原则。为什么可以？

        原因：SQL 的表名/字段名**不能**用 ? 占位符
        （标准 SQL 只允许对"值"用占位符，表名属于结构，不行）。
        所以只能用 f-string 拼进去。

        但这样一来就有 SQL 注入风险！所以做了双重防护：
            ① int(n)   ← 强制把 n 转成整数，恶意字符串转不了整数会直接报错
            ② [table_name] ← 用方括号包住表名，SQL Server 会把它当成
                             "普通标识符"，里面的特殊字符失去作用

    【新手警告】⚠ 即便如此，table_name 也**只应**来自程序内部
        （比如固定的表名列表），绝不能直接拿用户在界面上输入的文本！
        如果将来确实需要用户输入表名，必须做白名单校验
        （即：只允许事先允许的那几个表名，其他一律拒绝）。
    """
    # 【新手必学】f"..." 里的 {int(n)} 会先执行 int(n) 再插入字符串。
    # 比如 n="10" → int("10")=10 → 拼出来 "SELECT TOP 10 * FROM [表名]"
    sql = f"SELECT TOP {int(n)} * FROM [{table_name}]"
    # 因为值已经被拼进 SQL 了，所以这里没有占位符，params 是空列表。
    return sql, []