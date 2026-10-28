# -*- coding: utf-8 -*-
"""
【模块职责】主窗口 —— 整个程序的"总控台"

这是用户启动程序后看到的主界面。它把前面所有模块"串"起来：

    连接配置(ConnectionDialog) → 选表页(ReportSelectDialog)
        → 后台生成(GenerateWorker → ReportGenerator)
        → 显示进度/日志/勾稽核对结果

【新手须知】主窗口自己**不做**任何脏活累活：
    - 不直接连数据库（交给 core/db.py）
    - 不直接操作 Excel（交给 core/excel_engine.py 和 app/generator.py）
它只负责"接收用户点击 → 调用下层 → 把结果画到界面上"。
这种"界面层只管显示，业务层只管干活"的分工，是良好架构的核心思想。

【新手必学】三大 Qt 概念，本文件全部覆盖：
    1. 信号与槽（Signal / Slot）：Qt 的"事件订阅"机制，取代回调地狱
    2. QThread 后台线程：耗时任务放后台跑，否则界面会"假死"
    3. 布局（Layout）：控件不用手算坐标，交给 Layout 自动排列
"""
from __future__ import annotations

import os
import subprocess
from pathlib import Path

# 【新手须知】QThread 是"线程"，Signal 是"信号"。
# 两者都在 QtCore 里，不是 QtWidgets（Widgets 专管可见控件）。
from PySide6.QtCore import QThread, Signal
# 【新手必学】快捷键相关的两个类都在 QtGui 里，不在 QtWidgets：
#   QKeySequence → 把 "Ctrl+Shift+D" 这类文本解析成"按键组合"对象
#   QShortcut    → 把一个按键组合"绑定"到某个窗口函数上
from PySide6.QtGui import QKeySequence, QShortcut
from PySide6.QtWidgets import (
    QFileDialog,      # 文件/目录选择对话框
    QFormLayout,      # "标签: 控件" 两列式布局，做表单最好用
    QGroupBox,        # 带标题的边框分组框，用来给界面分区
    QHBoxLayout,      # 水平排列（left to right）
    QLabel,           # 纯文本标签
    QLineEdit,        # 单行文本输入框
    QMainWindow,      # 主窗口基类，自带菜单栏/状态栏/中央区域
    QMessageBox,      # 各种提示弹窗（警告/错误/询问）
    QPlainTextEdit,   # 多行纯文本框，适合当"日志显示区"
    QProgressBar,     # 进度条
    QPushButton,      # 按钮
    QSpinBox,         # 数字输入框（带上下箭头，防止非法输入）
    QVBoxLayout,      # 垂直排列（top to bottom）
    QWidget,          # 所有可视控件的"万能基类"，常用作容器
)

# 【新手须知】下面是本项目自己的模块。分成三组很好记：
#   app/    → 业务编排（干活的大脑）
#   config/ → 配置读写（记住用户设置）
#   core/   → 底层能力（日志、Excel、模板解析）
from app.generator import GenerateResult, ReportGenerator
from config.settings import AppConfig, ensure_dirs, save_config
from core.excel_engine import check_excel_available
from core.logging_setup import get_logger
from core.template import TemplateParser

# 【新手必学】每个模块顶部都这样取一个 logger。
# __name__ 会自动变成当前模块名（如 "ui.main_window"），
# 这样看日志时一眼就知道是哪一行代码打的。
logger = get_logger(__name__)


# ---------------------------------------------------------------------------
# 后台生成线程
# ---------------------------------------------------------------------------
class GenerateWorker(QThread):
    """在后台线程中执行报表生成。

    【新手难点】为什么需要线程？
        生成 40 多张报表可能耗时几十秒甚至几分钟。如果直接在界面的
        "主线程"里执行，界面就会卡住不动（专业名词叫"假死"），
        用户点哪都没反应，还可能被 Windows 标记为"未响应"。
        所以：耗时的活儿 → 丢到 QThread 后台线程去跑。

    【新手必学】QThread 的三个关键点：
        1. 把耗时逻辑写在 run() 方法里（不是 __init__！）
        2. 用 .start() 启动（不是 .run() 直接调用，
           直接调 run() 就变成"在当前线程执行"了，失去了多线程意义）
        3. 子线程**不能直接改界面**！只能发"信号"，
           让主线程收到信号后再去更新界面（本条规则极重要）
    """

    # 【新手必学】Signal(参数类型...) 定义一个"信号"。
    # 信号 = 广播电台：子线程只管"喊"，谁想听谁去 connect（订阅）。
    # 这里的 4 个参数类型分别是 int, int, str, str。
    progress = Signal(int, int, str, str)   # current, total, code, msg
    # object = 万能类型，这里装一个 GenerateResult 对象
    finished_ok = Signal(object)            # GenerateResult
    failed = Signal(str)                    # 错误信息文本

    def __init__(self, config: AppConfig, codes: list[str], parent=None):
        # 【新手须知】"子线程一定要写好 parent"是个好习惯：
        # parent 负责生命周期管理，父对象销毁时子对象会被一并回收，
        # 避免"野线程"还在跑、程序却退出了的尴尬。
        super().__init__(parent)
        self.config = config
        self.codes = codes

    def run(self) -> None:  # noqa: D102
        """线程启动后真正执行的代码（.start() 会来调它）。

        【新手难点】为什么整个 run() 要包在 try/except 里？
            因为子线程里一旦抛出未捕获的异常，Qt 只会打印一行警告，
            界面完全无感知，用户会以为"程序卡住了"。所以必须捕获，
            再通过 failed 信号把错误"喊"给主线程去弹窗。
        """
        try:
            # 【新手必学】把 self._on_progress 作为"回调函数"传给生成器：
            # 生成器每处理完一张表，就回头调一下这个函数报告进度。
            # 这就是"回调"（callback）——"你干你的，干到某一步记得叫我"。
            gen = ReportGenerator(self.config, progress=self._on_progress)
            result = gen.generate(self.codes)
            # 干完了，发信号通知主线程"成功了"
            self.finished_ok.emit(result)
        except Exception as exc:  # noqa: BLE001
            # exc_info=True 表示"把完整的异常堆栈也写进日志"，方便排错
            logger.error("生成线程异常：%s", exc, exc_info=True)
            self.failed.emit(str(exc))

    def _on_progress(self, current: int, total: int, code: str, msg: str) -> None:
        """把生成器报告来的进度，原样转发成 Qt 信号。

        【新手须知】这里为什么不直接更新进度条？
            因为本方法运行在**子线程**里，操作界面会崩溃。
            所以只能用 self.progress.emit(...) 发信号，
            由主线程的槽函数去更新（见 MainWindow._on_progress）。
        """
        self.progress.emit(current, total, code, msg)


# ---------------------------------------------------------------------------
# 主窗口
# ---------------------------------------------------------------------------
class MainWindow(QMainWindow):
    """报表生成主窗口。

    【新手必学】界面代码有个万能套路，记住它，一辈子受用：

        __init__():          设置标题尺寸 → 建界面 → 加载数据 → 刷新显示
        _build_ui():         只负责"把控件摆出来"，不做任何业务逻辑
        _on_xxx():           _on_ 开头的方法 = "某控件的点击/事件响应函数"
    """

    def __init__(self, config: AppConfig, parent=None):
        super().__init__(parent)
        self.config = config
        # 【新手须知】下划线开头（_catalog）表示"内部使用"，
        # 是 Python 的约定俗成：告诉别人"这是私有变量，别乱碰"。
        self._catalog = []          # 模板表页清单（SheetInfo 列表）
        self._worker = None         # 当前的后台线程（None 表示没有任务在跑）

        self.setWindowTitle("U8报表生成器")
        # 【新手须知】窗口高度比上一版（660）矮了不少。
        #   原因是【运行日志】默认收起，不再占垂直空间（详见 _build_ui）。
        #   客户看到的界面就是"账套与期间 / 输出 / 生成"三块，干净紧凑。
        #   维护者按 Ctrl+Shift+D 展开日志时，_toggle_dev_panel 会自动把窗口加高。
        self.resize(860, 470)

        # 【新手必学】初始化顺序很重要，一般是：
        #   先建界面 → 再加载数据 → 最后刷新显示
        # 如果顺序反了，界面控件可能还不存在，赋值的代码就会报错。
        self._build_ui()
        self._load_catalog()
        self._check_env()

    # ------------------------------------------------------------------
    # 构建界面
    # ------------------------------------------------------------------
    def _build_ui(self) -> None:
        """把界面上所有控件创建出来并摆好位置。

        【新手必学】Qt 三种核心布局容器（本方法三种都用了）：
            QVBoxLayout → 控件从上往下竖着排（"V" = Vertical）
            QHBoxLayout → 控件从左往右横着排（"H" = Horizontal）
            QFormLayout → "标签: 输入控件" 两列式表单，做设置页首选

        【新手难点】QMainWindow 不能直接把控件塞进去，必须：
            central = QWidget()              先造一个"中央容器"
            self.setCentralWidget(central)   把它设为主窗口内容
            root = QVBoxLayout(central)      再在它里面铺一层布局
        """
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)

        # ---- 账套与期间信息 ----
        # QGroupBox 用来给界面"分区"，视觉上更像一块一块的功能区
        #
        # 【新手难点】这里**故意不显示**"数据库连接串"和"模板路径"这两行。
        #   为什么不显示？它们属于"技术细节"：
        #     - 连接串又长又密（服务器\实例 / 库名 / 认证方式），普通用户看着发懵；
        #     - 模板路径往往长到半屏，还会把窗口顶得很宽。
        #   但信息并没有丢：
        #     - 连接信息随时可点"重新设置连接…"对话框查看和修改；
        #     - 模板是否解析成功，看下方【运行日志】里的提示即可。
        gb_conn = QGroupBox("账套与期间")
        form = QFormLayout(gb_conn)

        # 【新手必学】"self.xxx = 控件" 是固定套路。
        # 加 self. 是为了让别的函数也能访问到它（如 _on_click 里读它的值）。
        # 如果不加 self. 写成局部变量，函数一结束就被回收了，别的函数根本找不到。

        # 会计期间：年份 + 月份两个数字框，放在一个水平布局里
        row_period = QHBoxLayout()
        self.sp_year = QSpinBox()
        self.sp_year.setRange(1990, 2100)             # 限制范围，防止输入 99999
        self.sp_year.setValue(self.config.fiscal_year)  # 用配置里的值作为初值
        self.sp_month = QSpinBox()
        self.sp_month.setRange(1, 12)
        self.sp_month.setValue(self.config.fiscal_month)
        row_period.addWidget(QLabel("年"))
        row_period.addWidget(self.sp_year)
        row_period.addWidget(QLabel("月"))
        row_period.addWidget(self.sp_month)
        # 【新手必学】addStretch(1) = 塞一块"弹簧/空白"，
        # 把后面的东西挤到左边去。不加的话控件会被拉满整行，很难看。
        row_period.addStretch(1)
        form.addRow("会计期间：", row_period)

        self.ed_account = QLineEdit(self.config.account_set)  # 构造时传初值也可以在
        self.ed_account.setMaximumWidth(120)
        form.addRow("账套号：", self.ed_account)

        # 两个按钮放一行
        row_conn_btn = QHBoxLayout()
        self.btn_conn = QPushButton("重新设置连接…")
        # 【新手必学】.clicked.connect(func) = "这个按钮被点击时，去调 func"。
        # 注意 func 后面**没有括号**！有括号是"立刻调用"，没括号是"把函数本身交给 Qt"。
        self.btn_conn.clicked.connect(self._on_change_connection)
        row_conn_btn.addWidget(self.btn_conn)
        self.btn_reload_tpl = QPushButton("重新解析模板")
        self.btn_reload_tpl.clicked.connect(self._load_catalog)
        row_conn_btn.addWidget(self.btn_reload_tpl)
        row_conn_btn.addStretch(1)
        form.addRow("", row_conn_btn)   # 第一列留空，让按钮占满整行

        root.addWidget(gb_conn)   # 把"连接与账套"这整块塞进主布局

        # ---- 输出目录 ----
        gb_out = QGroupBox("输出")
        form_out = QFormLayout(gb_out)
        row_out = QHBoxLayout()
        self.ed_output = QLineEdit(self.config.output_dir)
        # 【新手必学】addWidget(控件, 1) 里的 1 是"伸缩因子"（stretch）。
        # 它让输入框占用多余的横向空间（被拉长），而旁边的浏览按钮保持原样。
        row_out.addWidget(self.ed_output, 1)
        self.btn_browse = QPushButton("浏览…")
        self.btn_browse.clicked.connect(self._on_browse_output)
        row_out.addWidget(self.btn_browse)
        form_out.addRow("目录：", row_out)
        root.addWidget(gb_out)

        # ---- 生成操作 ----
        gb_gen = QGroupBox("生成")
        v = QVBoxLayout(gb_gen)
        row_btn = QHBoxLayout()
        self.btn_select = QPushButton("选择表页并生成…")
        self.btn_select.clicked.connect(self._on_select_and_generate)
        row_btn.addWidget(self.btn_select)
        self.btn_generate_all = QPushButton("生成全部表页")
        self.btn_generate_all.clicked.connect(self._on_generate_all)
        row_btn.addWidget(self.btn_generate_all)
        row_btn.addStretch(1)
        v.addLayout(row_btn)   # 注意：布局里嵌布局用 addLayout，不是 addWidget

        self.progress = QProgressBar()
        # 【新手必学】"%v / %m" 是进度条的显示格式：
        #   %v = 当前值(value)   %m = 最大值(maximum)
        # 会显示成类似 "3 / 43" 的样子。
        self.progress.setFormat("%v / %m")
        v.addWidget(self.progress)

        self.lbl_progress = QLabel("就绪")
        v.addWidget(self.lbl_progress)
        root.addWidget(gb_gen)

        # ---- 结果区（客户可见）----
        # 【设计思路：把"客户要用的"和"维护要看的"彻底分开】
        #   客户真正关心的只有两件事：
        #     1. 这次到底成了没有、成了几张（→ lbl_result 一句话说清）
        #     2. 结果文件在哪、怎么打开（→ 下面两个按钮）
        #   至于逐张表的滚动日志、SQL 原文、审计文件路径，
        #   那是我（维护者）排错时才需要的，全部收进下面的"维护面板"。
        #   这样客户界面干净，也不会顺手点开一堆看不懂的英文字。
        gb_result = QGroupBox("结果")
        vres = QVBoxLayout(gb_result)
        self.lbl_result = QLabel("就绪")
        # 【新手必学】setWordWrap(True) = 文字太长时自动换行，
        # 而不是把窗口硬撑宽（否则用户一拉窗口，长句会把布局顶变形）。
        self.lbl_result.setWordWrap(True)
        vres.addWidget(self.lbl_result)

        row_tools = QHBoxLayout()
        self.btn_open_dir = QPushButton("打开输出目录")
        self.btn_open_dir.clicked.connect(self._on_open_output_dir)
        row_tools.addWidget(self.btn_open_dir)
        self.btn_open_file = QPushButton("打开结果文件")
        self.btn_open_file.clicked.connect(self._on_open_result_file)
        # 【新手必学】setEnabled(False) = 暂时"灰掉"按钮，点不动。
        # 因为一开始还没生成结果文件，让这个按钮能点反而会让用户困惑。
        # 等生成成功后再 setEnabled(True) 点亮它（见 _on_finished）。
        self.btn_open_file.setEnabled(False)
        row_tools.addWidget(self.btn_open_file)
        # addStretch(1) 在这里的作用是"弹簧"：把它左边的东西全推到左边，
        # 右边的控件（退出按钮）就自然贴到窗口最右侧，实现需求要求的"右下角"。
        row_tools.addStretch(1)

        # ---- 退出按钮（需求：界面右下角放一个退出按钮）----
        # 【设计要点：为什么不自己写一套退出逻辑，而是调 close()？】
        #   主窗口已经实现了 closeEvent()，里面处理了两件必须做的事：
        #     ① 若正在生成报表，先弹窗确认（避免留下损坏的半成品文件、
        #        或让 Excel 进程残留在后台）；
        #     ② 退出前保存期间/账套/输出目录等设置（下次启动能恢复）。
        #   若这里另写一套 os._exit() 或 sys.exit()，就会绕过这两件事：
        #   正在生成时被强退、设置也丢 —— 属于典型的"两套逻辑走岔"。
        #   所以按钮只负责"请求关闭窗口"，真正的收尾统一交给 closeEvent。
        self.btn_exit = QPushButton("退出")
        self.btn_exit.setToolTip("关闭程序（正在生成时会先征求确认）")
        self.btn_exit.clicked.connect(self.close)
        row_tools.addWidget(self.btn_exit)

        vres.addLayout(row_tools)
        root.addWidget(gb_result)

        # ---- 运行日志（维护面板：默认整块收起）----
        # 【新手难点】为什么整块隐藏，而不是只藏某个按钮？
        #   客户不需要看日志，而日志里天然含有大量内部信息
        #   （表名、SQL、驱动报错、路径……）。与其逐条去想"这句能不能给他看"，
        #   不如从根上解决：**日志区默认不存在于客户视野**。
        #   我按 Ctrl+Shift+D 才把它展开 —— 需要查错时才看，看完再按一次收起。
        #
        # 【重要：隐藏 ≠ 停止记录】
        #   控件 setVisible(False) 之后，往里面 appendPlainText 依然有效，
        #   只是不画在屏幕上。所以整个生成过程中的日志一行不落，
        #   我展开面板的瞬间就能看到完整历史，不用"从头再跑一遍"。
        self.gb_log = QGroupBox("运行日志（维护模式）")
        vlog = QVBoxLayout(self.gb_log)
        self.txt_log = QPlainTextEdit()
        self.txt_log.setReadOnly(True)   # 日志只读，防止用户乱改
        vlog.addWidget(self.txt_log)

        # SQL 审计入口只属于维护面板，跟着面板一起隐藏/显示。
        # 注意：这里**不再**单独 setVisible(False) 了 ——
        # 门已经被面板锁上，再给门里的抽屉上锁只会增加维护麻烦
        # （展开面板后还得手动找按钮）。这叫"只在必要层级做控制"。
        row_dev = QHBoxLayout()
        self.btn_open_audit = QPushButton("查看 SQL 审计")
        self.btn_open_audit.clicked.connect(self._on_open_audit)
        row_dev.addWidget(self.btn_open_audit)
        row_dev.addStretch(1)
        vlog.addLayout(row_dev)

        # 【新手必学】root.addWidget(控件, 1) 的 1 是"垂直伸缩因子"：
        # 让日志区吃掉下方所有多余空间，窗口拉大时日志框跟着变大。
        root.addWidget(self.gb_log, 1)
        # 默认收起。客户打开程序就是看不到它，界面干净紧凑。
        self.gb_log.setVisible(False)

        # 装好"维护快捷键"（详见 _install_dev_shortcut 的说明）
        self._install_dev_shortcut()

        # 记录最近一次生成的文件路径，None 表示"还没生成过"
        self._result_path: Path | None = None

    # ------------------------------------------------------------------
    # 初始化辅助
    # ------------------------------------------------------------------
    def _load_catalog(self) -> None:
        """解析模板文件，取出所有表页信息，存进 self._catalog。

        【新手必学】这种"可能失败的操作"标准写法：
            try:    正常流程
            except: 出事了怎么办
        日志里记一笔 + 弹窗告诉用户，让程序不至于直接崩溃退出。
        """
        try:
            parser = TemplateParser(self.config.template_path)
            parser.open()                 # 打开 Excel（只读解析）
            self._catalog = parser.sheets  # 拿到表页清单
            parser.close()                # 【新手须知】用完记得关，释放文件句柄
            # 【新手必学】列表推导式：[表达式 for 变量 in 列表 if 条件]
            # 这行等价于：遍历所有表页，只留下有 code（LK 编号）的，
            # 然后数一数有几个。if s.code 在 Python 里表示"非空即为真"。
            n = len([s for s in self._catalog if s.code])
            # 【新手必学】Path(...).name 只取"最后一段文件名"，
            # 例如 G:\...\模板.xls → 模板.xls。日志里不铺满整条长路径。
            self._log(f"模板解析成功：{Path(self.config.template_path).name}，正式报表 {n} 张。")
        except Exception as exc:  # noqa: BLE001
            self._catalog = []       # 失败就清空，防止拿着脏数据往下跑
            self._log(f"[错误] 模板解析失败：{exc}")
            QMessageBox.critical(self, "模板错误", f"模板解析失败：\n{exc}")

    def _check_env(self) -> None:
        """检查本机是否装了 Excel。

        【新手难点】为什么要查 Excel？
            本程序生成 .xls 文件靠的是 win32com（即"调用本机的真 Excel"
            来另存）。如果这台电脑没装 Excel，就生成不了文件。所以
            启动时先探一探，没装就提前警告用户，免得生成到一半才报错。
        """
        # 【新手必学】"元组解包"：函数返回 (ok, msg) 两个值，
        # 这一行就把它俩分别装进 ok 和 msg 两个变量里。
        ok, msg = check_excel_available()
        if ok:
            self._log(msg)
        else:
            self._log(f"[警告] {msg}")
            self._log("提示：缺少 Excel 环境时无法生成 .xls（格式依赖 win32com）。")
            # 【新手难点】这里为什么**还要**写 lbl_result？
            #   因为日志区默认对客户是隐藏的 —— 只写日志等于"没告诉他"。
            #   而"没装 Excel 就生成不了文件"是客户必须立刻知道的事，
            #   否则他会点半天"生成"却永远拿不到结果，还以为程序坏了。
            #   原则：**客户必须知情的事写 lbl_result，排错细节写日志**。
            self.lbl_result.setText(
                "环境提示：本机未检测到可用的 Excel，暂时无法生成 .xls 文件。"
            )

    def _log(self, text: str) -> None:
        """往日志框里追加一行文字（自己封装一下，省得每次写长长的控件名）。"""
        self.txt_log.appendPlainText(text)

    # ------------------------------------------------------------------
    # 事件
    # ------------------------------------------------------------------
    def _on_browse_output(self) -> None:
        """点击"浏览…"按钮：弹出文件夹选择框。

        【新手必学】QFileDialog.getExistingDirectory(父窗口, 标题, 起始目录)
        会返回用户选中的目录路径字符串；如果用户点了"取消"，
        返回空字符串 ""（空字符串在 if 里为假，所以会跳过赋值）。
        """
        d = QFileDialog.getExistingDirectory(self, "选择输出目录", self.ed_output.text())
        if d:
            self.ed_output.setText(d)

    def _on_change_connection(self) -> None:
        """点击"重新设置连接…"按钮：弹出连接对话框。

        【新手必学】这里用"函数内 import"（延迟导入），而不是写在文件顶部。
            好处：只有真正点开这个对话框时才会加载相关代码，
            启动更快，也避免模块之间的循环导入问题。
        """
        from ui.connection_dialog import ConnectionDialog

        dlg = ConnectionDialog(self.config.connection, self)
        # 【新手必学】dlg.exec() 会"阻塞"住，直到用户关闭对话框。
        # 用户点"确定"返回真值（True），点"取消"返回假值（False/0）。
        if dlg.exec():
            self.config.connection = dlg.result_config
            self._save_current_config()
            self._log("连接设置已更新并保存。")

    def _collect_period(self) -> None:
        """把界面上的期间/账套/输出目录，写回 config 对象（"界面 → 数据"）。

        【新手必学】这是"界面 → 模型"的收拢动作。为什么单独写一个方法？
            因为期间、账套、输出目录在很多地方都要用，写成一个方法后，
            生成前/退出前只要调一次，就能保证配置是最新的。
        """
        self.config.fiscal_year = self.sp_year.value()     # .value() 取数字框的值
        self.config.fiscal_month = self.sp_month.value()
        # 【新手必学】.strip() 去掉首尾空格；"A or B" 表示
        # "A 若为空则取 B"。防止用户把输入框清空后变成空字符串。
        self.config.account_set = self.ed_account.text().strip() or "001"
        self.config.output_dir = self.ed_output.text().strip() or self.config.output_dir

    def _save_current_config(self) -> None:
        """保存配置到磁盘（下次启动能记住上次的设置）。

        【新手难点】为什么这里只记 warning，不弹窗报错？
            因为"保存配置失败"是小事，不该吓到用户、更不该中断生成流程。
            所以降级处理：日志里提醒一下即可。
            （对比 _load_catalog 里的失败都要弹窗，因为那是致命问题。）
        """
        try:
            save_config(self.config)
        except Exception as exc:  # noqa: BLE001
            logger.warning("保存配置失败：%s", exc)

    def _on_select_and_generate(self) -> None:
        """点击"选择表页并生成…"：先弹选择框，再启动生成。"""
        # 【新手必学】"先判断前置条件，不满足就 early return（提前返回）"
        # 是极常用的写法，能避免一大串 if 缩进嵌套（"箭头型代码"）。
        if not self._ensure_catalog():
            return
        from ui.report_select_dialog import ReportSelectDialog

        dlg = ReportSelectDialog(
            self._catalog, self.config.last_selected_reports, self
        )
        # 【新手须知】"if not dlg.exec(): return" = 用户点了取消就结束。
        if not dlg.exec():
            return
        codes = dlg.selected_codes()
        # 记住这次的选择，下次打开选择框时默认勾上，省得用户重复勾
        self.config.last_selected_reports = codes
        self._start_generate(codes)

    def _on_generate_all(self) -> None:
        """点击"生成全部表页"：不弹选择框，全部表页都生成。"""
        if not self._ensure_catalog():
            return
        # 【新手必学】（续前文的列表推导式）
        # 遍历所有表页，取出其中有 code 的，组成一个纯 code 列表。
        codes = [s.code for s in self._catalog if s.code]
        self.config.last_selected_reports = codes
        self._start_generate(codes)

    def _ensure_catalog(self) -> bool:
        """检查模板是否解析成功。返回 True 表示可以继续生成。"""
        if not self._catalog:
            QMessageBox.warning(self, "提示", "模板尚未成功解析，无法生成。")
            return False
        return True

    # ------------------------------------------------------------------
    # 生成流程
    # ------------------------------------------------------------------
    def _start_generate(self, codes: list[str]) -> None:
        """启动生成任务（核心方法：准备 → 锁按钮 → 起线程）。

        【新手必学】这是全文件最值得学习的方法之一，注意看它做了三件事：
            ① 前置检查（防止重复点击、防止空列表）
            ② 更新界面状态（禁用按钮、重置进度条、写日志）
            ③ 创建并连接后台线程
        """
        # 【新手难点】为什么判断"线程在跑就拦截"？
        #   生成过程中若用户又点了生成，两个线程同时操作同一个 Excel，
        #   结果一定混乱甚至损坏文件。所以宁可弹窗提示，也不能放行。
        if self._worker is not None and self._worker.isRunning():
            QMessageBox.information(self, "提示", "正在生成中，请稍候…")
            return
        if not codes:
            QMessageBox.warning(self, "提示", "没有可生成的表页。")
            return

        self._collect_period()   # 把界面上的设置收进 config
        ensure_dirs()            # 确保 logs/ 等目录都存在
        # 【新手必学】把相对路径变成绝对路径（resolve 会解析成完整路径）。
        # 避免"从不同地方启动程序时，输出文件跑到不同位置"的诡异问题。
        self.config.output_dir = str(Path(self.config.output_dir).resolve())
        self._save_current_config()

        # 【新手必学】"开始耗时任务前，先把该禁用的按钮禁用掉"，
        # 等任务结束后再恢复（见 _restore_buttons）。这是界面防呆的标准做法。
        self.btn_select.setEnabled(False)
        self.btn_generate_all.setEnabled(False)
        self.btn_open_file.setEnabled(False)
        self.progress.setValue(0)
        # 进度条的最大值 = 要生成的表页数量（这样进度就是 3/43 这种）
        self.progress.setMaximum(len(codes))
        self._log(f"开始生成，共 {len(codes)} 张表页…")

        # 【新手必学】线程三连（记住这三步就够用）：
        #   ① 建线程对象
        #   ② 用 connect 把线程的信号连到主窗口的槽函数
        #   ③ .start() 启动（真正开始跑 run() 方法）
        self._worker = GenerateWorker(self.config, codes, self)
        self._worker.progress.connect(self._on_progress)
        self._worker.finished_ok.connect(self._on_finished)
        self._worker.failed.connect(self._on_failed)
        self._worker.start()

    def _on_progress(self, current: int, total: int, code: str, msg: str) -> None:
        """【槽函数】收到后台线程的进度信号，更新进度条与文字。

        注意：本方法运行在**主线程**（Qt 会跨线程投递信号，自动切回主线程），
        所以这里可以放心地操作界面控件。
        """
        self.progress.setMaximum(total)
        self.progress.setValue(current)
        self.lbl_progress.setText(msg)
        # 【设计要点：客户要知道"跑到哪了"，但不需要看技术细节】
        #   msg 形如"正在生成 LK01 资产负债表"，本身就够友好；
        #   这里再带上 current/total，客户一眼能判断还要等多久。
        if code:
            self.lbl_result.setText(f"正在生成（{current}/{total}）：{msg}")
        # 完整日志照旧往隐藏的日志框里堆，维护者展开即可看全过程。
        if code:
            self._log(msg)

    def _on_finished(self, result: GenerateResult) -> None:
        """【槽函数】收到"生成完成"信号，恢复界面并展示结果。

        【设计思路：两份报告，各给各看】
            - 客户：lbl_result 一句话（成没成、几成几败、下一步做什么）
            - 维护：txt_log 完整明细（每张表、耗时、取数统计、勾稽差异）
        """
        self._restore_buttons()

        # ---- 先判断：这次是不是"连文件都没产出"？----
        # 【新手难点：为什么必须单独处理这种情况？】
        #   有一种失败发生在"还没开始填表"之前：账套号+年度定位不到年度库。
        #   此时 result.success=False、output_path=None、一张表也没生成。
        #   如果沿用下面"成功 X 张 / 未成功 Y 张"的模板，就会显示成
        #   "生成结束：成功 0 张，未成功 0 张。结果文件已生成…"——
        #   前半句让人一头雾水，后半句更是**假的**（根本没有文件）。
        #   所以必须先分支：没有产出文件，就把"为什么没有"直接说给客户听。
        if (not result.output_path) and (not result.generated_codes):
            # 进度条**故意不拉满**：拉满会让客户以为"跑完了、成功了"。
            # 停在 0 + 一句明确的"未生成"，比一个满格进度条诚实得多。
            self.progress.setValue(0)
            self.lbl_progress.setText("未生成")
            # result.message 来自账套解析器，写的就是"人能照着做"的话
            # （如"未找到匹配的账套库：账套号 001、年度 2025。请核对…"），
            # 不含 SQL 原文与内部表名，可以直接给客户看。
            self.lbl_result.setText(f"未生成报表文件。\n{result.message}")
            self._log("-" * 40)
            self._log(result.message)
            self._log(f"[耗时] {result.elapsed:.2f} 秒")
            # 【为什么要再弹一次窗？】
            #   结果区那行字容易被忽略（客户这时往往正盯着进度条）。
            #   "一张表都没出来"是必须立刻知道的结论，弹窗能确保他看见，
            #   并且顺手把"下一步怎么做"讲清楚。
            #   result.message 是账套解析器写的中文提示，不含 SQL 与表名，
            #   可以直接展示（技术细节另有 logs 目录留档）。
            QMessageBox.warning(
                self, "未能生成报表",
                f"{result.message}\n\n"
                "本次没有产出任何文件。请按上面的提示核对后重新生成；"
                "如仍无法解决，请联系维护人员。",
            )
            return

        # 把进度条拉满，表示"全都跑完了"
        self.progress.setValue(self.progress.maximum())
        self.lbl_progress.setText(result.message)

        # ---- 客户区：一句话结论 ----
        # 【新手必学】用 len() 数"成功了几张"，用 len(result.failed) 数"失败几张"。
        #   这两个数字凑起来就是"客户唯一真正想知道的事"。
        n_ok = len(result.generated_codes)
        n_fail = len(result.failed)
        if n_fail == 0:
            self.lbl_result.setText(
                f"生成完成：成功 {n_ok} 张，用时 {result.elapsed:.1f} 秒。"
                "可点下方按钮打开结果文件。"
            )
        else:
            # 【新手难点】为什么不把失败原因直接写在这里？
            #   原因文字里难免带表名/驱动信息（属于内部细节），
            #   而且可能有十几行，塞进一句话会把界面撑爆。
            #   所以客户只需要知道"有几张没成、该找谁"，
            #   具体原因由维护者展开日志面板查看。
            self.lbl_result.setText(
                f"生成结束：成功 {n_ok} 张，未成功 {n_fail} 张。"
                "结果文件已生成，但部分表页数据不完整，请联系维护人员核查。"
            )

        self._log("-" * 40)          # 【新手须知】字符串乘数字 = 重复拼接，这里画出 40 个横线分隔
        self._log(result.message)

        # 【新手必学】result.failed 是个字典 {表页代码: 错误信息}。
        # "如果字典非空"就说明有失败的表页，逐个打印出来，方便用户查错。
        if result.failed:
            self._log("[失败表页]")
            for code, err in result.failed.items():
                self._log(f"  {code}: {err}")
            # 【新手必学：给用户"下一步该做什么"，而不是只报个错】
            #   普通财务人员看到"某某表失败"之后，最需要知道的是：
            #     1. 这跟我有没有关系（要不要我重新操作）；
            #     2. 真解决不了该找谁、该给谁看什么。
            #   所以这里给一句明确指引，具体的技术细节（SQL、表名、堆栈）
            #   已经完整写进 logs\ 目录下的文件，交给维护人员去查。
            self._log("  提示：以上为简要原因。"
                      "如无法自行解决，请把 logs 目录整体发给维护人员。")
        if result.audit_summary:
            self._log(f"[取数统计] {result.audit_summary}")
        # 【新手必学】:.2f 意思是"保留两位小数"，例如 12.3456 → 12.35
        self._log(f"[耗时] {result.elapsed:.2f} 秒")

        if result.output_path:
            self._result_path = Path(result.output_path)
            self.btn_open_file.setEnabled(True)   # 有结果文件了，点亮"打开"按钮
            self._log(f"[输出] {self._result_path}")

        # 勾稽关系核对报告
        if result.reconciliation_text:
            self._log("")
            self._log(result.reconciliation_text)

        if not result.success:
            QMessageBox.warning(self, "生成未完成", result.message)

    def _on_failed(self, err: str) -> None:
        """【槽函数】收到"失败"信号，恢复界面并弹窗提示错误。"""
        self._restore_buttons()
        self.lbl_progress.setText("生成失败")
        # 【设计要点：弹窗给客户看，日志给维护看】
        #   err 是整条流程崩掉的原始信息（可能很技术），
        #   直接弹给财务人员没有意义，还可能泄露内部结构。
        #   所以：客户区只说"失败了、请联系维护"，原始 err 只进日志。
        self.lbl_result.setText(
            "生成失败：本次未能产出结果文件。请联系维护人员处理。"
        )
        self._log(f"[错误] {err}")
        QMessageBox.critical(
            self, "生成失败",
            "本次未能完成生成。\n\n请联系维护人员处理，"
            "维护人员可展开日志区查看详细原因。",
        )

    def _restore_buttons(self) -> None:
        """把生成前禁用的按钮重新启用（生成结束/失败时都要调用）。

        【新手必学】把"恢复按钮"单独抽成一个方法很有必要：
            因为成功和失败两条路都要恢复按钮，写两遍就重复了。
            抽出来，两边都调一次即可 —— 这叫"消除重复代码"。
        """
        self.btn_select.setEnabled(True)
        self.btn_generate_all.setEnabled(True)

    # ------------------------------------------------------------------
    # 开发者入口（默认隐藏）
    # ------------------------------------------------------------------
    # 【新手必学】把"快捷键"定义成类常量，好处：
    #   1. 想换键位只改这一处；
    #   2. 注释能写清给谁用，日后接手的人不会一脸茫然。
    DEV_PANEL_SHORTCUT = "Ctrl+Shift+D"

    # 维护面板展开后的目标窗口高度（像素）。
    # 单独定义成常量，方便日后调整，不用在代码里翻那个魔法数字。
    DEV_PANEL_HEIGHT = 720

    def _install_dev_shortcut(self) -> None:
        """把"展开/收起维护面板"绑定到快捷键上。

        【新手难点】为什么一个快捷键能"一键切换"，而不是只"一键显示"？
            因为槽函数里判断当前状态再取反：
            现在藏着就显示，现在露着就收起 —— 同一个键来回切。
            比"一个键显示、另一个键隐藏"好记，也不容易忘。

        【新手难点】self._sc_panel 为什么要存成实例属性？
            QShortcut 对象一旦没人引用，会被 Python 的垃圾回收机制回收掉，
            快捷键随即失效（而且不报错，极难排查）。
            存到 self 上 = 让主窗口长期"牵着"它，保证一直有效。
        """
        self._sc_panel = QShortcut(QKeySequence(self.DEV_PANEL_SHORTCUT), self)
        # 【新手必学】QShortcut 的 activated 信号：按键被按下时触发。
        self._sc_panel.activated.connect(self._toggle_dev_panel)

    def _toggle_dev_panel(self) -> None:
        """展开 / 收起【运行日志（维护模式）】整块面板。

        【这个面板里有什么】
            - 逐张表的生成日志（成功/跳过/失败原因）
            - "查看 SQL 审计"按钮（打开 sql_audit.log，含原始 SQL）
            - 勾稽核对报告、耗时统计等技术信息
          这些对外是"内部细节"，对内是"排错必需品"，所以默认收起。
        """
        # 【新手难点】这里为什么用 isHidden()，而不是 isVisible()？
        #   isVisible() 判断的是"**此刻真的**显示在屏幕上没有"，
        #   它受"父窗口是否已显示"影响：如果主窗口还没 show()（例如程序
        #   刚启动、或自动化测试里没调 show()），子控件的 isVisible()
        #   一律返回 False —— 那我们写成 setVisible(not isVisible()) 就会
        #   变成"每次都设成 True"，切换逻辑直接失效。
        #   isHidden() 判断的是"控件**自身**有没有被显式隐藏"，与父窗口无关，
        #   正是这里需要的语义。
        #
        #   所以：isHidden() 为 True（现在藏着）→ setVisible(True) 展开
        #         isHidden() 为 False（现在露着）→ setVisible(False) 收起
        showing = self.gb_log.isHidden()
        self.gb_log.setVisible(showing)

        # 【新手必学：顺手给用户一个"够用"的窗口高度】
        #   面板收起时窗口只有几百像素高，突然展开日志区会把下面的内容
        #   挤得几乎看不见。所以展开时若窗口偏矮，就自动加高到 DEV_PANEL_HEIGHT。
        #   用 max(...) 保证"只加高、不压缩"：如果我本来就手动拉得比它高，
        #   就不要把我拉的尺寸改小。
        if showing:
            self.resize(self.width(), max(self.height(), self.DEV_PANEL_HEIGHT))
        # 【重要：故意不写日志】
        #   如果在这里 _log("再按 Ctrl+Shift+D 可收起")，
        #   就等于把"维护快捷键"直接印在日志里 —— 虽然日志本身也藏起来了，
        #   但万一误开就会泄露。所以保持静默：知道这个键的人本来就懂。

    # ------------------------------------------------------------------
    # 工具按钮
    # ------------------------------------------------------------------
    def _on_open_output_dir(self) -> None:
        """打开输出目录（不存在就先创建）。"""
        d = Path(self.ed_output.text().strip() or self.config.output_dir)
        # 【新手必学】mkdir(parents=True, exist_ok=True)：
        #   parents=True  → 父目录不存在也一并创建（相当于 mkdir -p）
        #   exist_ok=True → 目录已存在也不报错
        d.mkdir(parents=True, exist_ok=True)
        self._open_path(d)

    def _on_open_result_file(self) -> None:
        """打开最近生成的结果 Excel 文件。"""
        # 【新手必学】"if 值 and 条件:" 这种短路写法：
        # 先确定 _result_path 不是 None，再调 .exists()，
        # 否则对 None 调方法会直接抛异常。
        if self._result_path and self._result_path.exists():
            self._open_path(self._result_path)
        else:
            QMessageBox.information(self, "提示", "尚无结果文件。")

    def _on_open_audit(self) -> None:
        """打开 SQL 审计日志（用来核查"程序到底取了哪些数"）。"""
        from config.settings import PROJECT_ROOT
        # 【新手必学】Path 对象可以用 "/" 运算符拼路径！
        # PROJECT_ROOT / "logs" / "sql_audit.log" 等价于拼接成
        # 带平台的正确路径，比手写字符串加反斜杠更安全。
        audit = PROJECT_ROOT / "logs" / "sql_audit.log"
        if audit.exists():
            self._open_path(audit)
        else:
            QMessageBox.information(self, "提示", f"审计日志尚未生成：\n{audit}")

    def _open_path(self, path: Path) -> None:
        """用系统默认程序打开文件或目录。

        【新手难点】为什么有两层 try？
            Windows 上 os.startfile 最省事（双击效果），但它只在 Windows 有。
            万一失败，就退回用 explorer（资源管理器）打开。
            两层都失败，才弹窗告诉用户。这叫"降级方案"，能提高健壮性。
        """
        try:
            os.startfile(str(path))  # type: ignore[attr-defined]
        except Exception:
            try:
                subprocess.Popen(["explorer", str(path)])
            except Exception as exc:  # noqa: BLE001
                QMessageBox.warning(self, "打开失败", f"{path}\n{exc}")

    # ------------------------------------------------------------------
    def closeEvent(self, event):  # noqa: N802
        """【Qt 生命周期钩子】用户点右上角 × 关闭窗口时会自动调用本方法。

        【新手必学】closeEvent 是 Qt 预留的"回调钩子"：
            你不需要手动调用它，Qt 在关闭窗口那一刻会自动来调。
            名字必须**一模一样**（拼错就失效了），所以文件顶部加了
            "from __future__ import annotations"，避免一些兼容问题。

        【新手难点】为什么这里要问一句"确定要退出吗？"
            生成过程会占用 Excel 和数据库连接。若跑到一半用户强退，
            可能留下损坏的半成品文件、或让 Excel 进程残留在后台。
            所以正在生成时，必须让用户确认。
        """
        if self._worker is not None and self._worker.isRunning():
            ret = QMessageBox.question(
                self, "确认退出", "正在生成报表，确定要退出吗？",
            )
            # 【新手必学】event.ignore() = "忽略这次关闭请求"，
            # 窗口就不会关掉了。反过来 event.accept() 才是同意关闭。
            if ret != QMessageBox.Yes:
                event.ignore()
                return
        # 退出前把最后的设置保存一下，下次启动就能恢复
        self._collect_period()
        self._save_current_config()
        super().closeEvent(event)   # 交给父类做真正关闭窗口的收尾工作