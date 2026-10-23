# -*- coding: utf-8 -*-
"""
报表生成编排器
============================================================

【模块职责】
    "编排"（orchestrate）= 把各个零件按正确顺序组装起来跑一遍。
    本模块是**整个生成流程的总指挥**，负责：
    1. 打开模板（只读）。
    2. 仅保留用户选中的表页（说明类页一并保留）。
    3. 按 LK 代码顺序，逐表 fetch → write → verify。
    4. 另存为一个多表页文件。
    5. 生成完成后做勾稽关系核对。

【设计要点（为什么这么设计）】
    - 顺序生表：符合勾稽依赖，进度可见。
    - 单一 DB 连接 + 共享缓存：避免重复查询（性能关键）。
    - 单表失败不影响其他表：捕获异常并记录，继续下一张。
    - 通过 progress_callback 向界面报告进度（界面才能显示进度条）。

【架构位置】
    本模块属于"应用服务层"，处在中间：
        上层 UI 调用它 → 它调用 core/（数据库、Excel）+ reports/（各表逻辑）
    UI 不需要知道"表怎么生成"的细节，只调用 generate() 拿结果。
    这叫"关注点分离"。
"""
from __future__ import annotations

import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Optional

from config.settings import AppConfig
from core.context import Period, ReportContext
from core.db import Database, _brief_error
from core.excel_engine import ExcelEngine, ExcelEngineError
from core.logging_setup import get_logger
from core.template import SheetInfo, TemplateParser
from reports import registry

logger = get_logger(__name__)

# =====================================================================
# 【新手必学：Callable 类型提示】
#   Callable[[int, int, str, str], None] 表示"一个函数"：
#       - 接收 4 个参数，类型分别是 int、int、str、str
#       - 返回 None（即不返回值）
#   这就像给"回调函数"定了个规格说明书。
#
# 【什么是"回调"（callback）？】
#   生成报表时，界面想知道"现在生成到第几张了"，以便更新进度条。
#   但 generator 又不该直接操作界面（否则分层就白分了）。
#   解决办法：让界面"传一个函数进来"，generator 在关键点调用它。
#   这个被传入的函数就叫"回调"。
#
#   进度回调签名：(当前序号, 总数, 当前代码, 消息)
# =====================================================================
ProgressCallback = Callable[[int, int, str, str], None]


@dataclass
class GenerateResult:
    """
    生成结果（打包返回给界面）。

    【新手必学：为什么用"结果对象"而不是返回一堆值？】
        如果写成 return success, path, codes, failed, elapsed, message...
        调用方要记顺序、容易接错。打包成一个对象后：
            result.success
            result.output_path
        既清晰又不易错。

    【新手难点：field(default_factory=list) 的原因】
        不能写 generated_codes: list[str] = []！
        那样所有 GenerateResult 会共享同一个列表（经典 Python 陷阱）。
        必须用 field(default_factory=list)，让每个对象各自新建一个空列表。
        （详见 core/context.py 里对 default_factory 的讲解）
    """

    success: bool = False
    output_path: Optional[Path] = None
    generated_codes: list[str] = field(default_factory=list)
    failed: dict[str, str] = field(default_factory=dict)  # code -> 错误信息
    elapsed: float = 0.0
    message: str = ""
    audit_summary: str = ""
    reconciliation_text: str = ""   # 勾稽关系核对报告


class ReportGenerator:
    """
    报表生成编排器。

    【典型用法】
        gen = ReportGenerator(config, progress=我的进度函数)
        result = gen.generate(["LK01", "LK02"])
    """

    def __init__(self, config: AppConfig, progress: Optional[ProgressCallback] = None):
        """
        参数：
            config  : 应用配置（含连接信息、模板路径、期间、输出目录等）。
            progress: 可选的进度回调函数。不传（None）就静默运行。
        """
        self.config = config
        self.progress = progress

    def _report(self, current: int, total: int, code: str, msg: str) -> None:
        """
        内部辅助：安全地调用进度回调。

        【新手须知：为什么要判断 if self.progress？】
            如果不判断，而调用方没传回调（None），
            就会报 "NoneType object is not callable"。
            这里先确认"有回调才调用"，程序更健壮。
            这种"跳过可选功能"的写法非常常见。
        """
        if self.progress:
            self.progress(current, total, code, msg)

    # ------------------------------------------------------------------
    # 主流程
    # ------------------------------------------------------------------
    def generate(self, selected_codes: list[str]) -> GenerateResult:
        """
        生成选定表页。

        参数 selected_codes：用户选中的表页代码列表（按预期顺序）。

        返回：GenerateResult 对象。

        【流程总览（建议对照代码一步步读）】
            1. 解析模板目录 → 拿到"代码→名称"的对照
            2. 建数据库连接 + 建上下文（缓存容器）
            3. 加载已实现的报表模块（触发注册）
            4. 构造报表实例列表
            5. 打开 Excel，只保留选中的表页
            6. 按顺序：取数 → 填表 → 核对登记
            7. 另存为新文件
            8. 做勾稽关系核对
        """
        start = time.perf_counter()   # 记录总开始时间（用于算总耗时）
        result = GenerateResult()     # 准备结果对象

        # ===== 第1步：解析模板目录（获取代码→名称）=====
        parser = TemplateParser(self.config.template_path)
        parser.open()
        catalog: list[SheetInfo] = parser.sheets   # 全部表页信息（含说明类）
        parser.close()

        # 【新手须知：提前返回（early return）】
        #   没选任何表页就没必要继续了，直接返回。
        #   这种"尽早退出"能减少缩进层级，让主流程更清晰。
        if not selected_codes:
            result.message = "未选择任何表页"
            return result

        # ===== 第2步：建立数据库连接与上下文 =====
        db = Database(self.config.connection)
        period = Period(self.config.fiscal_year, self.config.fiscal_month)

        try:
            db.connect()
        except Exception as exc:
            # 【设计选择】允许无数据库时生成"骨架报表"（用于调试模板复制）。
            # 好处：开发期没配数据库也能验证"表页复制、另存"是否正常。
            # 注意：此时各报表的取数会失败，但框架不会崩。
            logger.warning("数据库连接失败，将生成空报表：%s", exc)

        # 上下文：贯穿本次生成，负责缓存与审计（详见 core/context.py）
        ctx = ReportContext(db, period, self.config.account_set)

        # ===== 第3步：加载内置报表模块 =====
        # 【新手难点】这行只是在"import 模块"，但它会触发各报表类上的
        # @register 装饰器执行，从而完成注册。详见 reports/registry.py。
        registry.load_builtin_reports()

        # ===== 第4步：构造报表实例（按选中顺序）=====
        # 已实现的用真类，未实现的自动用占位类，所以这里一定有 43 个对象可用。
        reports = registry.build_report_set(selected_codes, catalog)

        # ===== 第5步：准备输出路径，打开 Excel =====
        out_dir = Path(self.config.output_dir)
        # parents=True 自动建多级目录，exist_ok=True 已存在也不报错
        out_dir.mkdir(parents=True, exist_ok=True)

        # 用时间戳命名，避免多次生成互相覆盖
        stamp = time.strftime("%Y%m%d_%H%M%S")
        out_path = out_dir / f"财务报表_{period.key}_{stamp}.xls"

        # 【新手必学：try/except/finally 三件套】
        #   try     : 正常要做的活
        #   except  : 出特定错误时的补救（这里只处理 Excel 引擎错误）
        #   finally : 无论成败都执行（这里保证数据库连接一定关闭）
        try:
            # with 会在退出时自动 close Excel（并保证不保存模板）
            with ExcelEngine(self.config.template_path) as engine:
                self._report(0, len(reports), "", "正在复制选定表页…")
                # 只保留选中的表页（其余报表页删掉，说明类页保留）
                engine.keep_only(selected_codes)

                # ===== 第6步：顺序生成每一张表 =====
                total = len(reports)
                # 【新手必学：enumerate(..., start=1)】
                #   让序号从 1 开始（而不是默认的 0），更符合"第 1 张、第 2 张"的说法。
                for i, rep in enumerate(reports, start=1):
                    self._report(i, total, rep.code, f"正在生成 {rep.code} {rep.name}")
                    try:
                        hook = engine.sheet(rep.code)   # 拿到该页的写入句柄
                        rep.generate(hook, ctx)          # 取数→填表→核对登记
                        result.generated_codes.append(rep.code)
                    except Exception as exc:
                        # 【关键设计：单表失败不影响其他表】
                        #   记录错误、跳过本张、继续下一张。
                        #   43 张表里坏 1 张，不该让另外 42 张也白干。
                        #   exc_info=True 会把完整调用栈写进日志，便于定位。
                        #
                        # 【重要：界面只放"简化后的原因"】
                        #   完整原始异常只进 logs/app.log（下面这行 logger.error）；
                        #   而 result.failed 里的文字最终会显示在界面【运行日志】上，
                        #   所以必须先过一道 _brief_error，去掉表名、SQL、驱动噪声。
                        logger.error("[%s] 生成失败：%s", rep.code, exc, exc_info=True)
                        result.failed[rep.code] = _brief_error(exc)

                # ===== 第7步：另存为新文件 =====
                self._report(total, total, "", "正在保存文件…")
                engine.save_as(out_path, fmt="xls")

            # with 块结束后，Excel 已关闭。下面开始整理结果。
            result.success = True
            result.output_path = out_path
            result.message = f"生成完成，共 {len(result.generated_codes)} 张报表"

            # ===== 第8步：勾稽关系核对 =====
            # 基于各表 verify() 登记的关键指标（存于 ctx.report_data）做交叉验证。
            try:
                # 【新手须知：函数内导入（延迟导入）】
                #   把 import 写在函数里，好处是"用到才加载"，
                #   既能加快启动，也能避免"循环导入"问题。
                from verify.reconciliation import Reconciler
                result.reconciliation_text = Reconciler(ctx).report_text()
            except Exception as exc:  # noqa: BLE001
                # 核对失败不该影响"报表已成功生成"这个结论，
                # 所以只记警告，不改变 result.success。
                logger.warning("勾稽核对失败：%s", exc)

        except ExcelEngineError as exc:
            # 只捕获"Excel 引擎"这一类错误（比如没装 Excel、保存失败）
            # 【重要】message 会显示在界面与弹窗上，所以同样走一遍简化，
            #        完整原文仍由下面这行 logger.error 留在 app.log 里。
            result.message = f"Excel 引擎错误：{_brief_error(exc)}"
            logger.error("Excel 引擎错误：%s", exc, exc_info=True)
        finally:
            # 【重要】无论成功失败，都要关闭数据库连接，避免连接泄漏。
            db.close()

        # 收尾：记录总耗时与取数审计摘要
        result.elapsed = time.perf_counter() - start
        result.audit_summary = ctx.audit_summary()
        logger.info("生成结束：%s，耗时 %.2fs", result.message, result.elapsed)
        return result