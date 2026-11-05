# -*- coding: utf-8 -*-
"""账套 → 业务库 解析逻辑的单元测试。

运行方式（在项目根目录执行）：
    python -m unittest discover -s tests -v

【本文件测什么？】
    core/u8_account.py 的三条判定规则（20261105 现场口径）：

        情形一：该账套只有 1 条登记记录
                → 本年、前一年都用这唯一库

        情形二：查询年度**恰好等于**某条记录的 iBeginYear
                → 本年用该条；前一年用"相邻较小"那条；
                  若该条是最小（首年）→ 无前一年数据

        情形三：查询年度与所有 iBeginYear 都不相等
                → 本年是"iBeginYear 较小者"那条；
                  前一年**也是同一条**

【为什么不连真数据库？】
    判定逻辑是**纯计算**：给一组 (库名, 起始年度) 记录 + 一个查询年度，
    必然得到唯一答案。用假 Database 就能在毫秒级验证十几个分支，
    且不依赖任何 U8 环境。真实连接的事，由手工端到端验证负责。
"""
from __future__ import annotations

import sys
import unittest
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from core.db import DatabaseError      # noqa: E402
from core.u8_account import (          # noqa: E402
    AccountResolveError,
    resolve_account_database,
)


class FakeAccountDB:
    """假数据库：只实现 resolve 用到的那三个方法。

    【为什么这样够用？】
        resolve_account_database 只做三件事：
            ① db.query(...)             取登记记录
            ② db.database_exists(...)   验证库存在
            ③ db.switch_database(...)   切换连接
        把这三件"假装"掉，判定逻辑就能独立测试。
    """

    def __init__(self, rows, existing=None, fail_query=False):
        """
        参数：
            rows       : 模拟 UA_AccountDatabase 的查询结果
                         [(cDatabase, iBeginYear, iEndYear), ...]
            existing   : 服务器上真实存在的库名集合；None 表示"全部存在"
            fail_query : True 时 db.query 抛异常，用于测权限错误分支
        """
        self._rows = rows
        self._existing = existing
        self._fail = fail_query
        self.switched_to = None
        self.queries: list = []

    def query(self, sql, params=None, database=None, tag=""):
        self.queries.append((sql, params))
        if self._fail:
            # 【必须抛 DatabaseError 而非 RuntimeError】
            #   被测试的解析器只捕获 DatabaseError（那是 db.query 的
            #   真实异常类型）。若这里抛 RuntimeError，就等于模拟了一种
            #   现实中不会发生的情形，测不出我们想验的那条分支。
            raise DatabaseError("模拟：无系统库查询权限")
        # 只对"查登记记录"的那条 SQL 返回数据；
        # 辅助提示（list_accounts）返回空，避免干扰
        if "UA_AccountDatabase" in sql and "WHERE cAcc_Id" in sql:
            return list(self._rows)
        return []

    def database_exists(self, database):
        if self._existing is None:
            return True
        return database in self._existing

    def switch_database(self, database):
        self.switched_to = database


def resolve(rows, year, existing=None, switch=True):
    """跑一次解析，返回 (target, fake_db)。"""
    fake = FakeAccountDB(rows, existing=existing)
    target = resolve_account_database(
        fake, "001", year, system_database="UFSystem", switch=switch)
    return target, fake


class TestSingleRecord(unittest.TestCase):
    """情形一：该账套只有 1 条登记记录（跨年度单库）。"""

    def test_single_record_current_and_prev_same(self):
        """唯一记录 2013→UFDATA_001，查 2026：本年与上年都是这个库。

        【这是最常见的现场形态】
            U8 推荐做法就是"一个账套一个库，从建立年度一直用下去"。
            若此处判断错（例如旧代码以为"仅 2013 年有效"），
            客户选 2026 年会直接报"未找到匹配的账套库"。
        """
        target, fake = resolve(
            [("UFDATA_001", 2013, None)], year=2026)

        self.assertEqual(target.database, "UFDATA_001")
        self.assertEqual(target.prev_database, "UFDATA_001")
        self.assertEqual(target.prev_year, 2025)
        self.assertTrue(target.has_prev)
        self.assertTrue(target.is_single_database)
        # 连接应被切到该库
        self.assertEqual(fake.switched_to, "UFDATA_001")

    def test_single_record_earliest_year_ok(self):
        """查的正好是起始年（2013）也应成功，且上年同库。"""
        target, _ = resolve([("UFDATA_001", 2013, None)], year=2013)
        self.assertEqual(target.database, "UFDATA_001")
        self.assertEqual(target.prev_database, "UFDATA_001")

    def test_single_record_year_before_begin_raises(self):
        """查 2012（早于账套起始年 2013）必须报错，不能猜库。

        【为什么不"就近取那个库"？】
            2012 年该账套根本不存在，硬取 2013 的库会拿到一份
            "看起来正常但与年度不符"的数据 —— 比报错危险得多。
        """
        with self.assertRaises(AccountResolveError) as ctx:
            resolve([("UFDATA_001", 2013, None)], year=2012)
        self.assertIn("2013", str(ctx.exception))
        self.assertIn("早于", str(ctx.exception))


class TestExactYearMatch(unittest.TestCase):
    """情形二：查询年度恰好等于某条记录的 iBeginYear。"""

    def setUp(self):
        # 三条分年度库：2013→A、2015→B、2018→C
        self.rows = [
            ("UFDATA_001_2013", 2013, 2013),
            ("UFDATA_001_2015", 2015, 2015),
            ("UFDATA_001_2018", 2018, 2018),
        ]

    def test_exact_middle_year(self):
        """查 2015：本年=B，上年取相邻较小者=2013 的 A。"""
        target, _ = resolve(self.rows, year=2015)
        self.assertEqual(target.database, "UFDATA_001_2015")
        self.assertEqual(target.prev_database, "UFDATA_001_2013")
        self.assertEqual(target.prev_year, 2014)
        self.assertTrue(target.prev_exact)

    def test_exact_last_year(self):
        """查 2018：本年=C，上年=2015 的 B。"""
        target, _ = resolve(self.rows, year=2018)
        self.assertEqual(target.database, "UFDATA_001_2018")
        self.assertEqual(target.prev_database, "UFDATA_001_2015")

    def test_exact_first_year_has_no_prev(self):
        """查 2013（首年）：本年=A，**没有前一年数据**。

        【为什么必须为空而不是"也返回 A"？】
            2013 是该账套建账第一年，2012 年不存在任何数据。
            若返回 A，上层会去 A 库里查 2012 年凭证，结果为空，
            然后填一堆 0 到"上年同期"列 —— 看起来像"去年一分钱没花"。
            prev_database=None 能让上层明确知道"这一列应当留空"。
        """
        target, _ = resolve(self.rows, year=2013)
        self.assertEqual(target.database, "UFDATA_001_2013")
        self.assertIsNone(target.prev_database)
        self.assertIsNone(target.prev_year)
        self.assertFalse(target.has_prev)


class TestBetweenYears(unittest.TestCase):
    """情形三：查询年度与所有 iBeginYear 都不相等（落在区间内）。"""

    def setUp(self):
        self.rows = [
            ("UFDATA_001_2013", 2013, 2013),
            ("UFDATA_001_2015", 2015, 2015),
            ("UFDATA_001_2018", 2018, 2018),
        ]

    def test_between_2015_and_2018(self):
        """查 2016：落在 2015 与 2018 之间 → 本年=B，前一年**也是 B**。

        【这是最容易写错的一条】
            直觉容易想成"前一年去 2015 那条找"，但规则明确写的是
            "查询的前一年数据也同在这个库"。
            因为该库覆盖 2015 到 2018 之前的所有年度，
            2015 与 2016 的数据物理上都在 B 库里。
        """
        target, _ = resolve(self.rows, year=2016)
        self.assertEqual(target.database, "UFDATA_001_2015")
        self.assertEqual(target.prev_database, "UFDATA_001_2015",
                         "前一年应与本年同库（情形三规则）")
        self.assertFalse(target.prev_exact,
                         "非精确匹配时 prev_exact 应为 False")

    def test_between_2017(self):
        target, _ = resolve(self.rows, year=2017)
        self.assertEqual(target.database, "UFDATA_001_2015")
        self.assertEqual(target.prev_database, "UFDATA_001_2015")

    def test_after_last_record(self):
        """查 2020：大于最大记录 2018 → 本年=C，前一年也是 C。"""
        target, _ = resolve(self.rows, year=2020)
        self.assertEqual(target.database, "UFDATA_001_2018")
        self.assertEqual(target.prev_database, "UFDATA_001_2018")

    def test_before_first_record_raises(self):
        """查 2012（早于最小 2013）→ 报错。"""
        with self.assertRaises(AccountResolveError):
            resolve(self.rows, year=2012)


class TestEdgeCases(unittest.TestCase):
    """边界与异常分支。"""

    def test_no_rows_raises(self):
        """该账套在登记表里一条记录都没有 → 报错并提示核对。"""
        with self.assertRaises(AccountResolveError) as ctx:
            resolve([], year=2026)
        self.assertIn("未找到账套", str(ctx.exception))

    def test_library_not_exists_raises(self):
        """登记了库名、但服务器上没有这个库 → 报错并指出库名。"""
        with self.assertRaises(AccountResolveError) as ctx:
            resolve([("UFDATA_001", 2013, None)], year=2026,
                    existing={"UFDATA_999"})
        msg = str(ctx.exception)
        self.assertIn("UFDATA_001", msg)
        self.assertIn("并不存在", msg)

    def test_prev_library_not_exists_raises(self):
        """本年库存在、但前一年库缺失 → 也要报错。

        【为什么前一年库也要验证？】
            分年度多库模式下，前一年库可能已被删除或迁移。
            若不验证，生成时会在"取上年同期数"这一步失败，
            错误信息会指向某个具体报表，掩盖真正原因。
        """
        rows = [("UFDATA_001_2013", 2013, 2013),
                ("UFDATA_001_2015", 2015, 2015)]
        with self.assertRaises(AccountResolveError) as ctx:
            resolve(rows, year=2015, existing={"UFDATA_001_2015"})
        msg = str(ctx.exception)
        self.assertIn("前一年度", msg)
        self.assertIn("UFDATA_001_2013", msg)

    def test_dirty_rows_skipped(self):
        """库名为空 / 起始年度为 NULL 的脏记录应被跳过，不影响正常判定。"""
        rows = [
            ("", 2013, None),                 # 库名空 → 跳过
            ("UFDATA_001_2019", None, None),  # 起始年度空 → 跳过
            ("UFDATA_001", 2013, None),       # 有效
        ]
        target, _ = resolve(rows, year=2026)
        self.assertEqual(target.database, "UFDATA_001")
        self.assertEqual(len(target.entries), 1)

    def test_query_failure_translated(self):
        """系统库查询失败 → 转成可读提示，不把 SQL 错误抛给界面。"""
        fake = FakeAccountDB([], fail_query=True)
        with self.assertRaises(AccountResolveError) as ctx:
            resolve_account_database(fake, "001", 2026)
        self.assertIn("无法读取", str(ctx.exception))
        self.assertIn("权限", str(ctx.exception))

    def test_account_set_normalized(self):
        """账套号 "1" 应被规范化为 "001" 再去查。"""
        fake = FakeAccountDB([("UFDATA_001", 2013, None)])
        resolve_account_database(fake, "1", 2026)
        _, params = fake.queries[0]
        self.assertEqual(params, ["001"])

    def test_switch_false_does_not_switch(self):
        """switch=False 时不应切换连接（供"仅预览"场景使用）。"""
        _, fake = resolve([("UFDATA_001", 2013, None)], year=2026, switch=False)
        self.assertIsNone(fake.switched_to)

    def test_describe_mentions_both(self):
        """describe() 应同时提到本年库与前一年库，便于日志核对。"""
        target, _ = resolve(
            [("UFDATA_001_2013", 2013, 2013),
             ("UFDATA_001_2015", 2015, 2015)], year=2015)
        text = target.describe()
        self.assertIn("UFDATA_001_2015", text)
        self.assertIn("UFDATA_001_2013", text)


class TestSqlBuilder(unittest.TestCase):
    """校验 account_databases() 生成的 SQL 与参数。"""

    def test_sql_has_no_year_filter(self):
        """SQL 不应按年度过滤 —— 必须取回该账套全部记录。

        【为什么单独锁这条？】
            早期版本在 SQL 里写了年度判断条件，
            结果每次只能拿到一条记录，前一年库无从推导。
            这条测试确保：只要有人重新加上年度过滤，立刻失败。
        """
        from queries import u8
        sql, params = u8.account_databases("UFSystem", "001")

        self.assertEqual(params, ["001"],
                         "只应传账套号一个参数")
        self.assertNotIn("BETWEEN", sql.upper(),
                         "不应出现年度区间过滤")
        self.assertNotIn("iEndYear IS NULL", sql,
                         "不应再依赖 iEndYear 判断（那是错误口径）")
        self.assertIn("ORDER BY iBeginYear ASC", sql,
                      "必须按起始年度升序返回，供上层推导相邻关系")
        self.assertIn("UA_AccountDatabase", sql)

    def test_sql_quotes_system_database(self):
        """系统库名必须用 [] 包裹（标识符不能走占位符）。"""
        from queries import u8
        sql, _ = u8.account_databases("UFSystem", "001")
        self.assertIn("[UFSystem]", sql)


if __name__ == "__main__":
    unittest.main(verbosity=2)