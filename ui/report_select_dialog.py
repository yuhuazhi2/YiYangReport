# -*- coding: utf-8 -*-
"""
报表表页选择对话框
============================================================

【设计目标（对应需求）】
    1. 一次交互完成"要生成哪些表页"的选择，避免多轮对话。
    2. 列出模板中全部正式报表（LK 代码 + 名称），复选框选择。
    3. 支持全选 / 全不选 / 反选 / 关键字筛选。
    4. 记住上次选择，下次默认勾选。

输出：选中的表页代码列表（按模板顺序）。

【新手必学：QListWidget 的"带复选框列表"怎么用？】
    Qt 的列表控件不是天生带复选框，需要两步：
      1. 创建 QListWidgetItem（列表项）。
      2. item.setCheckState(Qt.Checked) 让它显示复选框。
    读取用户是否勾选，就用 item.checkState()。
"""
from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QDialog,
    QDialogButtonBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QPushButton,
    QVBoxLayout,
)

from core.logging_setup import get_logger
from core.template import SheetInfo

logger = get_logger(__name__)


class ReportSelectDialog(QDialog):
    """选择要生成的报表表页"""

    def __init__(self, sheets: list[SheetInfo], checked_codes=None, parent=None):
        """
        参数：
            sheets        : 模板全部表页（含说明类页）。
            checked_codes : 上次选中的代码（用于回填勾选状态）。可为 None。
            parent        : 父窗口。
        """
        super().__init__(parent)
        self.setWindowTitle("选择要生成的报表")
        self.resize(560, 620)   # 宽 560、高 620

        # 仅保留正式报表（有 LK 代码的），说明类页不参与勾选
        # 【新手必学：列表推导式做过滤】只留 s.code 非空（即真）的项
        self._sheets = [s for s in sheets if s.code]

        # 【新手必学：set 集合的使用】
        #   checked_codes or []  → 若为 None 就用空列表，避免报错
        #   set(...)             → 转成集合，查找"某个代码是否在已选里"更快
        self._checked = set(checked_codes or [])

        self._build_ui()
        self._populate()

    # ------------------------------------------------------------------
    def _build_ui(self) -> None:
        """创建控件与布局。"""
        root = QVBoxLayout(self)

        # 顶部说明
        tip = QLabel(
            "勾选需要生成的报表表页。生成的文件将只包含所选表页（外加说明类页）。\n"
            "可按需全选或只选部分；上次选择会被记住。"
        )
        tip.setWordWrap(True)
        root.addWidget(tip)

        # ---- 筛选框 ----
        row_filter = QHBoxLayout()
        row_filter.addWidget(QLabel("筛选："))
        self.ed_filter = QLineEdit()
        self.ed_filter.setPlaceholderText("输入 LK 代码或名称关键字")
        # 【信号槽】每次文字变化都调用 _apply_filter，
        # 从而实现"边打字边过滤"（即时搜索）的效果。
        self.ed_filter.textChanged.connect(self._apply_filter)
        # 参数 1 是伸缩因子，让输入框占满右侧空间
        row_filter.addWidget(self.ed_filter, 1)
        root.addLayout(row_filter)

        # ---- 列表 ----
        self.list = QListWidget()
        # 【新手须知：NoSelection 的用意】
        #   这里用复选框来"选中"，不需要行高亮的那种选中效果（容易混淆），
        #   所以禁掉默认的 SelectionMode。
        self.list.setSelectionMode(QAbstractItemView.NoSelection)
        # 隔行不同底色，长列表更易读
        self.list.setAlternatingRowColors(True)
        # 参数 1 作为伸缩因子，让列表占满剩余的纵向空间
        root.addWidget(self.list, 1)

        # ---- 操作按钮（全选/全不选/反选）----
        row_ops = QHBoxLayout()
        self.btn_all = QPushButton("全选")
        self.btn_none = QPushButton("全不选")
        self.btn_invert = QPushButton("反选")
        # 【新手必学：lambda 传递参数的技巧】
        #   clicked 信号不带参数（或带的参数我们不要），
        #   但 _set_all 需要一个 bool 参数。
        #   于是用 lambda: self._set_all(True) 包一层，
        #   相当于"创建一个无参的小函数，内部固定调用 _set_all(True)"。
        self.btn_all.clicked.connect(lambda: self._set_all(True))
        self.btn_none.clicked.connect(lambda: self._set_all(False))
        self.btn_invert.clicked.connect(self._invert)
        row_ops.addWidget(self.btn_all)
        row_ops.addWidget(self.btn_none)
        row_ops.addWidget(self.btn_invert)
        # addStretch(1) 插入一个"弹性空白"，把后面的内容推到右边
        row_ops.addStretch(1)
        self.lbl_count = QLabel("")
        row_ops.addWidget(self.lbl_count)
        root.addLayout(row_ops)

        # ---- 确定/取消 ----
        self.buttons = QDialogButtonBox()
        self.btn_ok = self.buttons.addButton("确定生成", QDialogButtonBox.AcceptRole)
        self.btn_cancel = self.buttons.addButton("取消", QDialogButtonBox.RejectRole)
        self.btn_ok.clicked.connect(self._on_ok)
        self.btn_cancel.clicked.connect(self.reject)
        root.addWidget(self.buttons)

    # ------------------------------------------------------------------
    def _populate(self) -> None:
        """把表页清单填充到列表控件里。"""
        self.list.clear()

        # 【逻辑说明】
        #   上次有选择记录(_checked 非空) → 按上次的选择回填勾选
        #   上次没有记录                  → 默认全部勾选（方便第一次使用）
        default_checked = bool(self._checked)

        for s in self._sheets:
            # 列表项显示"代码 + 名称"（display_name 是 SheetInfo 的属性）
            item = QListWidgetItem(s.display_name)

            # 【新手必学：setData 存隐藏数据】
            #   把表页代码藏在列表项里（界面不显示），
            #   后面需要时用 item.data(Qt.UserRole) 取回来。
            #   Qt.UserRole 是 Qt 预留的"用户自定义数据"槽位，随便用。
            item.setData(Qt.UserRole, s.code)

            if default_checked:
                # 【新手必学：三目表达式决定勾选状态】
                item.setCheckState(
                    Qt.Checked if s.code in self._checked else Qt.Unchecked
                )
            else:
                item.setCheckState(Qt.Checked)

            # 【新手必学：用位运算"追加"一个标志位】
            #   item.flags() 拿到当前允许的操作（如可点、可拖拽等），
            #   | Qt.ItemIsUserCheckable 用按位或"补上"可勾选这一项。
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)

            self.list.addItem(item)

        # 【新手必学：lambda *_: ... 忽略参数的写法】
        #   itemChanged 信号会传一个 item 参数，但我们不用它，
        #   于是用 *_ 接收并忽略，同时每次变化都刷新计数显示。
        self.list.itemChanged.connect(lambda *_: self._update_count())
        self._update_count()   # 初始刷新一次

    def _apply_filter(self, text: str) -> None:
        """
        按关键字隐藏不匹配的列表项。

        【新手必学：隐藏而不是删除】
            用 item.setHidden(True/False) 而不是删除项，
            这样清空筛选框后能立刻恢复显示，且不会丢失勾选状态。
        """
        text = text.strip().upper()
        for i in range(self.list.count()):
            item = self.list.item(i)
            if not text:
                item.setHidden(False)   # 没有筛选条件，全部显示
                continue
            label = item.text().upper()
            code = str(item.data(Qt.UserRole) or "").upper()
            # 【逻辑】名称和代码里都找不到关键字 → 隐藏
            #   and 表示"两个都不包含"才隐藏（两者任一匹配就显示）
            item.setHidden(text not in label and text not in code)

    def _set_all(self, checked: bool) -> None:
        """
        全选 / 全不选。

        【重要细节】只操作"当前可见"的项（!item.isHidden()）。
            这样配合筛选功能，用户能实现"筛选出某类表后整批勾选"的便捷操作。
            —— 这正是需求里"选中部分表页"的实用场景。
        """
        state = Qt.Checked if checked else Qt.Unchecked
        for i in range(self.list.count()):
            item = self.list.item(i)
            if not item.isHidden():
                item.setCheckState(state)
        self._update_count()

    def _invert(self) -> None:
        """反选：已勾选的取消，未勾选的勾上（同样只针对可见项）。"""
        for i in range(self.list.count()):
            item = self.list.item(i)
            if item.isHidden():
                continue
            # 【新手必学：三目表达式做"切换"】
            #   当前是 Checked 就变 Unchecked，否则变 Checked。
            item.setCheckState(
                Qt.Unchecked if item.checkState() == Qt.Checked else Qt.Checked
            )
        self._update_count()

    def _update_count(self) -> None:
        """刷新"已选 N / 总数"的计数文字。"""
        n = len(self.selected_codes())
        self.lbl_count.setText(f"已选 {n} / {len(self._sheets)} 张")

    # ------------------------------------------------------------------
    def selected_codes(self) -> list[str]:
        """
        返回选中代码（保持模板顺序）。

        【为什么"保持模板顺序"很重要？】
            报表之间有勾稽依赖，按模板(LK)顺序生成，前面的表先取数入缓存，
            后面的表就能复用，既快又符合业务逻辑。
            因为我们是按列表顺序遍历的，所以天然就是模板顺序。
        """
        codes = []
        for i in range(self.list.count()):
            item = self.list.item(i)
            if item.checkState() == Qt.Checked:
                # 把存在 UserRole 里的代码取出来
                codes.append(str(item.data(Qt.UserRole)))
        return codes

    def _on_ok(self) -> None:
        """确定：至少选一张才允许关闭。"""
        if not self.selected_codes():
            # 【新手须知：函数内导入】
            #   这里才用到 QMessageBox，写在函数里可以少一次启动时的加载。
            from PySide6.QtWidgets import QMessageBox
            QMessageBox.warning(self, "提示", "请至少选择一张报表。")
            return
        self.accept()