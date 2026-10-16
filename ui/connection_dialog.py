# -*- coding: utf-8 -*-
"""
数据库连接对话框
============================================================

【设计目标（对应需求）】
    1. 程序启动后首先弹出本对话框，录入 SQL Server 连接信息。
    2. 一次录入，测试连接成功后保存，下次启动自动回填（无需重复输入）。
    3. 连接成功后可选数据库（U8 年度库，如 UFDATA_001_2026）。

本对话框只负责"获取并验证连接配置"，不负责生成报表
（职责单一，符合分层设计原则）。

【新手必学：Qt 对话框（QDialog）的基本套路】
    1. 继承 QDialog。
    2. 在 __init__ 里创建各种控件（QLineEdit、QPushButton...）。
    3. 用"布局"（Layout）把控件排好位置。
    4. 用"信号槽"（signal/slot）把"按钮点击"和"处理函数"连起来。
    5. 用户点"确定"调 accept()，点"取消"调 reject()。
    6. 调用方用 dlg.exec() 弹出它，这是"阻塞式"的——会停在这里等用户操作完。

【新手必学：什么是"信号槽"？】
    Qt 的核心机制。控件在特定时刻会发出"信号"（如 clicked 被点击），
    你可以把这个信号连到自己的函数（叫"槽"）上：
        self.btn_test.clicked.connect(self._on_test)
    含义：用户点这个按钮时，自动调用 self._on_test()。
    它类似网页里的"事件监听"（addEventListener）。
"""
from __future__ import annotations

from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QSpinBox,
    QVBoxLayout,
)

from config.settings import ConnectionConfig
from core.db import Database
from core.logging_setup import get_logger

logger = get_logger(__name__)

# 【设计说明：ODBC 驱动不再让用户选择】
#   本工具的使用者是普通财务人员，不懂"ODBC 驱动"为何物。
#   因此界面上不再出现驱动下拉框，驱动连接由 Database.connect()
#   自动按优先级依次尝试本机已安装的驱动（详见 core/db.py）。


class ConnectionDialog(QDialog):
    """SQL Server 连接设置对话框"""

    def __init__(self, config: ConnectionConfig, parent=None):
        """
        参数：
            config: 已有的连接配置（用于回填上次的输入）。
            parent: 父窗口。传给 super().__init__ 后，
                   对话框会跟着父窗口居中、并在任务栏正确归组。
        """
        super().__init__(parent)
        self.setWindowTitle("U8报表生成器")
        self.setMinimumWidth(520)

        self._config = config    # 保存一份配置（点确定时会更新它）
        self._db = None          # 测试连接成功后，把 Database 对象暂存起来

        # 【新手必学：把界面搭建拆成独立方法】
        #   构造函数只负责"编排"，具体控件创建放到 _build_ui 里。
        #   这样代码短小、易读，也方便单独维护。
        self._build_ui()
        self._load_from_config(config)   # 回填上次的输入

    # ------------------------------------------------------------------
    # 构建界面
    # ------------------------------------------------------------------
    def _build_ui(self) -> None:
        """
        创建所有控件并排版。

        【新手必学：布局（Layout）是怎么工作的？】
            Qt 不用"坐标"摆控件，而用"布局"自动排列：
              - QVBoxLayout : 竖直方向，从上到下依次排
              - QHBoxLayout : 水平方向，从左到右依次排
              - QFormLayout : 左侧标签 + 右侧输入框，两列对齐（最适合表单）
            addWidget(控件) 添加一个控件，addRow("标签", 控件) 加一行表单。
            好处：窗口大小变化时控件会自动重新排布，不用手算坐标。
        """
        # 根布局：整体从上到下（竖直）
        root = QVBoxLayout(self)

        # ---- 顶部提示文字 ----
        tip = QLabel(
            "请输入用友 U8 后台 SQL Server 的连接信息。\n"
            "连接成功并确定后，本机将记住这些信息，下次启动自动回填。\n"
            "（数据库驱动由程序自动识别，无需您设置。）"
        )
        tip.setWordWrap(True)   # 允许自动换行，避免文字被截断
        root.addWidget(tip)

        # ---- 服务器信息（用 QGroupBox 分组，带边框和标题，更清晰）----
        gb_server = QGroupBox("服务器")
        form_server = QFormLayout(gb_server)   # 把布局装进分组框里

        self.ed_server = QLineEdit()
        # setPlaceholderText：输入框为空时显示的灰色提示文字
        self.ed_server.setPlaceholderText("本机可填 localhost；局域网填服务器 IP，如 192.168.1.10")
        form_server.addRow("服务器地址：", self.ed_server)

        self.ed_instance = QLineEdit()
        self.ed_instance.setPlaceholderText("默认实例留空；命名实例如 SQLEXPRESS")
        form_server.addRow("实例名：", self.ed_instance)

        self.sp_port = QSpinBox()          # 数字输入框
        self.sp_port.setRange(0, 65535)    # 端口范围
        self.sp_port.setSpecialValueText("默认")  # 值为 0 时显示"默认"而不是 0
        self.sp_port.setValue(1433)        # SQL Server 默认端口
        form_server.addRow("端口：", self.sp_port)

        root.addWidget(gb_server)

        # ---- 身份验证 ----
        gb_auth = QGroupBox("身份验证")
        form_auth = QFormLayout(gb_auth)

        self.chk_win = QCheckBox("使用 Windows 集成认证（当前 Windows 登录账户）")
        # 【信号槽示例】复选框状态变化时，调用 _on_auth_changed 联动禁用/启用账号密码框
        self.chk_win.stateChanged.connect(self._on_auth_changed)
        form_auth.addRow(self.chk_win)

        self.ed_user = QLineEdit()
        form_auth.addRow("登录名：", self.ed_user)

        self.ed_pwd = QLineEdit()
        # 【安全】密码框设为 Password 模式，输入内容显示为圆点，防止旁人看到
        self.ed_pwd.setEchoMode(QLineEdit.Password)
        form_auth.addRow("密码：", self.ed_pwd)

        root.addWidget(gb_auth)

        # ---- 数据库 ----
        gb_db = QGroupBox("数据库")
        form_db = QFormLayout(gb_db)

        # 数据库那一行需要"下拉框 + 按钮"横向并排，所以套一个水平布局
        row_db = QHBoxLayout()
        self.cmb_db = QComboBox()
        self.cmb_db.setEditable(True)      # 允许手输（列表中没有的也能填）
        self.cmb_db.setToolTip("如 UFDATA_001_2026")
        # addWidget(控件, 1) 里的 1 是"伸缩因子"（stretch），
        # 表示这个控件占据多余空间的 1 份，从而"撑开"占满宽度。
        row_db.addWidget(self.cmb_db, 1)

        self.btn_load_db = QPushButton("载入库列表")
        self.btn_load_db.setToolTip("连接服务器后列出可用数据库")
        self.btn_load_db.clicked.connect(self._on_load_databases)
        row_db.addWidget(self.btn_load_db)
        form_db.addRow("数据库：", row_db)

        # 【注意】此处不再放置"ODBC 驱动"下拉框：
        #   驱动由程序自动依次尝试，用户无需（也无法）选择。

        root.addWidget(gb_db)

        # ---- 状态提示文字（显示"连接成功/失败"等信息）----
        self.lbl_status = QLabel("")
        self.lbl_status.setWordWrap(True)
        root.addWidget(self.lbl_status)

        # ---- 底部按钮组 ----
        # QDialogButtonBox 是 Qt 提供的"标准按钮条"，
        # 用 ActionRole/AcceptRole/RejectRole 声明按钮的"角色"，
        # Qt 会自动处理确定/取消的标准行为，界面也更符合系统习惯。
        self.buttons = QDialogButtonBox()
        self.btn_test = self.buttons.addButton("测试连接", QDialogButtonBox.ActionRole)
        self.btn_ok = self.buttons.addButton("确定", QDialogButtonBox.AcceptRole)
        self.btn_cancel = self.buttons.addButton("取消", QDialogButtonBox.RejectRole)

        self.btn_test.clicked.connect(self._on_test)
        self.btn_ok.clicked.connect(self._on_ok)
        # 【新手须知：reject 是 QDialog 自带的槽函数】
        #   调用它会关闭对话框并返回"取消"结果，无需自己写。
        self.btn_cancel.clicked.connect(self.reject)
        root.addWidget(self.buttons)

    # ------------------------------------------------------------------
    # 数据绑定（配置 <-> 界面 互相同步）
    # ------------------------------------------------------------------
    def _load_from_config(self, c: ConnectionConfig) -> None:
        """
        把配置对象里的值"回填"到界面控件上。

        【新手须知：or 的妙用】
            c.server or "localhost"
            如果 c.server 为空（空字符串在布尔判断里是假），
            就用 "localhost" 作为默认值。这是很常用的简写方式。
        """
        self.ed_server.setText(c.server or "localhost")
        self.ed_instance.setText(c.instance)
        self.sp_port.setValue(c.port)
        self.chk_win.setChecked(c.use_windows_auth)
        self.ed_user.setText(c.username)
        self.ed_pwd.setText(c.password)
        self.cmb_db.setCurrentText(c.database)
        # 驱动不再从界面回填（界面已无此项），完全交给程序自动识别。
        self._on_auth_changed()   # 手动触发一次，让账号密码框的启用状态正确

    def _on_auth_changed(self, *_args) -> None:
        """
        Windows 认证勾选状态改变时：联动启用/禁用账号密码框。

        【新手必学：(*_args) 的写法与用途】
            信号可能带参数（比如 stateChanged 会传新状态值），
            但我们不关心它。写 *_args 表示"接收任意数量的参数，但忽略它们"。
            如果直接写 ()，某些信号因参数不匹配会报错。
            【原理】参数前的 * 表示"可变位置参数"，会把收到的参数打包成元组。
        """
        win = self.chk_win.isChecked()
        # Windows 认证时不需要账号密码，禁用它们（灰掉）
        self.ed_user.setEnabled(not win)
        self.ed_pwd.setEnabled(not win)

    def collect_config(self) -> ConnectionConfig:
        """
        从界面收集连接配置，打包成一个 ConnectionConfig 对象。

        【新手须知：.strip() 清空首尾空白】
            用户输入时可能不小心多打空格，strip() 去掉首尾空白，
            避免"看着一样却连不上"的问题。
        """
        c = ConnectionConfig()
        c.server = self.ed_server.text().strip() or "localhost"
        c.instance = self.ed_instance.text().strip()
        c.port = self.sp_port.value()
        c.use_windows_auth = self.chk_win.isChecked()
        c.username = self.ed_user.text().strip()
        # 【注意】密码不 strip()，因为空格可能是密码的真实一部分
        c.password = self.ed_pwd.text()
        c.database = self.cmb_db.currentText().strip()
        # 【关键】驱动一律留空，表示"自动"：
        #   Database.connect() 会自动挑一个可用的 SQL Server 驱动，
        #   连通后把真实驱动名回填到配置并保存，下次即可直接复用。
        c.driver = ""
        return c

    # ------------------------------------------------------------------
    # 交互
    # ------------------------------------------------------------------
    def _set_status(self, text: str, ok=None) -> None:
        """
        在状态标签上显示文字，并按结果设置颜色。

        参数 ok：
            True  → 绿色（成功）
            False → 红色（失败）
            None  → 黑色（进行中/普通提示）

        【新手须知：为什么要传 None？】
            表示"还不知道结果"（比如"正在连接…"时），
            这时用默认黑色，不误导用户。
        """
        color = "#000000"
        if ok is True:
            color = "#1a7f37"    # 绿色
        elif ok is False:
            color = "#c00000"    # 红色
        # setStyleSheet 用类似 CSS 的语法设置控件外观
        self.lbl_status.setStyleSheet(f"color:{color};")
        self.lbl_status.setText(text)

    def _on_test(self) -> None:
        """
        "测试连接"按钮的槽函数。

        【新手难点：为什么用 try/finally 恢复按钮状态？】
            连接 DB 可能耗时几秒甚至超时。若期间按钮仍可点，
            用户可能连点多次，开出很多连接。
            所以先禁用按钮，无论成功失败最后都用 finally 恢复，
            保证按钮不会"永久变灰"。
        """
        cfg = self.collect_config()
        self._set_status("正在连接…", None)
        self.btn_test.setEnabled(False)
        try:
            db = Database(cfg)
            ok, msg = db.test_connection()   # 返回 (是否成功, 消息)
            if ok:
                self._db = db                # 成功则暂存连接对象
                self._set_status(msg, True)
            else:
                self._set_status(msg, False)
        finally:
            self.btn_test.setEnabled(True)   # 无论如何都恢复按钮

    def _on_load_databases(self) -> None:
        """
        "载入库列表"按钮：连接服务器后列出所有数据库，供用户选择。

        【需求对应】用户拿到一台陌生的 U8 服务器时，未必记得年度库名。
                   点这个按钮就能列出全部库，并自动优先选中 UFDATA 开头的。
        """
        cfg = self.collect_config()
        self.btn_load_db.setEnabled(False)
        self._set_status("正在连接并读取数据库列表…", None)
        try:
            db = Database(cfg)
            ok, msg = db.test_connection()
            if not ok:
                self._set_status(msg, False)
                return

            names = db.list_databases()
            current = self.cmb_db.currentText()

            self.cmb_db.clear()
            self.cmb_db.addItems(names)

            # 优先定位 U8 年度库（用友的库名通常以 UFDATA 开头）
            # 【新手必学：列表推导式 + 字符串方法】
            #   [n for n in names if n.upper().startswith("UFDATA")]
            #   筛出所有以 UFDATA 开头的库名（先转大写再判断，避免大小写问题）。
            u8 = [n for n in names if n.upper().startswith("UFDATA")]
            if u8:
                # 取最后一个（通常是年度最靠后的库，如 UFDATA_001_2026）
                self.cmb_db.setCurrentText(u8[-1])
            elif current:
                self.cmb_db.setCurrentText(current)   # 保留用户原来的输入

            db.close()   # 用完就关，不占连接
            self._set_status(f"读取到 {len(names)} 个数据库，U8 年度库 {len(u8)} 个。", True)
        except Exception as exc:  # noqa: BLE001
            # 【新手须知：捕获宽泛异常】
            #   这里故意捕获所有异常并显示给用户，
            #   因为"列库失败"的原因很多（网络、权限、驱动…），
            #   与其逐一列举，不如统一提示。noqa 注释告诉检查工具这没问题。
            self._set_status(f"读取数据库列表失败：{exc}", False)
        finally:
            self.btn_load_db.setEnabled(True)

    def _on_ok(self) -> None:
        """
        "确定"按钮：先验证连接可用，通过后才接受对话框。

        【设计要点】不在"测试连接"成功时就允许关闭，
                  而是确定时再验证一次，
                  确保"用户点确定时，配置确实是能连通的"，
                  避免保存了错误的配置。
        """
        cfg = self.collect_config()

        # 必填校验：数据库不能为空
        if not cfg.database:
            QMessageBox.warning(self, "提示", "请选择或输入要连接的数据库（U8 年度库）。")
            return   # return 直接结束函数，对话框保持打开

        ok, msg = Database(cfg).test_connection()
        if not ok:
            QMessageBox.critical(self, "连接失败", msg)
            return

        self._config = cfg   # 验证通过，更新配置
        # 【新手必学：accept() 与 reject()】
        #   accept() → 关闭对话框并返回"接受"（exec() 返回 QDialog.Accepted）
        #   reject() → 关闭对话框并返回"拒绝"
        #   调用方据此判断用户是"确定"还是"取消"。
        self.accept()

    # ------------------------------------------------------------------
    @property
    def result_config(self) -> ConnectionConfig:
        """
        对话框被"接受"后的连接配置。

        调用方用法：
            if dlg.exec() == QDialog.Accepted:
                cfg = dlg.result_config
        """
        return self._config