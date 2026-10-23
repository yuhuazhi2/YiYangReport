# -*- coding: utf-8 -*-
"""
报表表页选择对话框
============================================================

【设计目标（对应需求）】
    1. 一次交互完成"要生成哪些表页"的选择，避免多轮对话。
    2. 列出模板中全部正式报表（LK 代码 + 名称），复选框选择。
    3. 支持全选 / 全不选 / 反选 / 关键字筛选。
    4. 【常用 9 表】用深蓝色加粗 + ★ 星号突出显示，
       并提供"只看常用9表 / 全表浏览"一键切换。
    5. 处于"只看常用9表"时，选择范围就收窄为这 9 张：
       9 表以外的表即使之前被勾选，也**不计入已选、不会被生成**，
       计数显示为"已选 9 / 43"（分母恒为模板总表数 43）。
    6. 记住上次选择，下次默认勾选。

输出：选中的表页代码列表（按模板顺序）。

【新手必学：QListWidget 的"带复选框列表"怎么用？】
    Qt 的列表控件不是天生带复选框，需要两步：
      1. 创建 QListWidgetItem（列表项）。
      2. item.setCheckState(Qt.Checked) 让它显示复选框。
    读取用户是否勾选，就用 item.checkState()。
"""
from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor
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

from config.settings import COMMON_REPORT_CODES
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

        # 常用 9 表也存成集合。
        # 【新手必学：为什么用 set 而不是 list？】
        #   判断"某代码是不是常用表"（in 运算）时，
        #   set 是"哈希查找"，几乎瞬间完成；list 要逐个比对。
        #   列表长时差别明显，这是很常用的优化小技巧。
        self._common_codes = set(COMMON_REPORT_CODES)

        # 【状态变量】当前是否处于"只看常用9表"模式。
        #   False = 显示全部表（默认）；True = 只显示常用 9 张。
        #   按钮点击时会翻转这个值，从而在两种模式间切换。
        self._common_only = False

        self._build_ui()
        self._populate()

    # ------------------------------------------------------------------
    def _build_ui(self) -> None:
        """创建控件与布局。"""
        root = QVBoxLayout(self)

        # 顶部说明
        tip = QLabel(
            "勾选需要生成的报表表页。生成的文件将只包含所选表页（外加说明类页）。\n"
            "带 ★ 的深蓝色加粗项为常用报表；点下方“只看常用9表”后，"
            "只有这 9 张会参与生成，其它表不计入已选。"
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

        # 【新增】常用表切换按钮。
        #   初始文字是"只看常用9表"：点击后列表只留 9 张常用表；
        #   进入该模式后按钮文字会变成"全表浏览"，再点一次恢复全部。
        #   让按钮文字随状态变化，用户一眼就知道"再点会发生什么"。
        self.btn_common = QPushButton("只看常用9表")
        self.btn_common.setToolTip(
            "只显示日常最常用的 9 张报表"
            "（LK01/LK02/LK03/LK25/LK26/LK27/LK28/LK31/LK33）"
        )
        self.btn_common.clicked.connect(self._toggle_common_only)

        row_ops.addWidget(self.btn_all)
        row_ops.addWidget(self.btn_none)
        row_ops.addWidget(self.btn_invert)
        row_ops.addWidget(self.btn_common)
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
            is_common = s.code in self._common_codes

            # 常用表在文字前面加一个 ★，即使不看颜色也能一眼认出
            #   （照顾色弱用户，也方便截图/打印时辨认）。
            text = ("★ " + s.display_name) if is_common else s.display_name
            item = QListWidgetItem(text)

            # 【常用表用特别颜色显示】
            #   setForeground(...) 设置文字颜色；
            #   QColor("#1565c0") 是深蓝色，与普通黑字对比明显。
            #   item.font() 取出当前字体，改成加粗后再 setFont 写回去。
            if is_common:
                item.setForeground(QColor("#1565c0"))
                f = item.font()
                f.setBold(True)
                item.setFont(f)

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

    def _apply_filter(self, text: str = "") -> None:
        """
        按关键字 +“只看常用9表”开关，共同决定每一项显示还是隐藏。

        【新手必学：两个条件要"同时满足"】
            一项必须同时通过两道关卡才显示：
              1) 若打开了常用模式，它必须属于常用 9 表；
              2) 若输入了关键字，它的代码或名称必须包含该关键字。
            任一道不通过就隐藏——所以代码里用 continue 直接跳去处理下一项。

        【新手必学：隐藏而不是删除】
            用 item.setHidden(True/False) 而不是删除项，
            这样清空筛选框、或切回"全表浏览"后能立刻恢复显示，
            而且不会丢失已有的勾选状态。
        """
        keyword = text.strip().upper()
        for i in range(self.list.count()):
            item = self.list.item(i)
            code = str(item.data(Qt.UserRole) or "").upper()

            # 关卡1：常用模式下，凡不是常用表的，一律隐藏
            if self._common_only and code not in self._common_codes:
                item.setHidden(True)
                continue

            # 关卡2：没有关键字 → 全部显示
            if not keyword:
                item.setHidden(False)
                continue

            label = item.text().upper()
            # 名称和代码里都找不到关键字 → 隐藏
            #   and 表示"两个都不包含"才隐藏（任一匹配就显示）
            item.setHidden(keyword not in label and keyword not in code)

    def _toggle_common_only(self) -> None:
        """
        在“只看常用9表”与“全表浏览”两种模式之间切换。

        【新手必学：开关型按钮的标准三步】
            1. 翻转状态变量 _common_only；
            2. 根据新状态更新按钮文字（告诉用户"再点会怎样"）；
            3. 重新执行一次过滤，让列表立刻反映变化。
            这三步几乎是所有"切换"按钮的通用写法，可以套用到别处。
        """
        self._common_only = not self._common_only

        if self._common_only:
            # 已进入常用模式：按钮现在的作用是"退回全部"，文字要跟着改
            self.btn_common.setText("全表浏览")
            self.btn_common.setToolTip("恢复显示模板中的全部报表表页")
        else:
            # 已回到全表模式：按钮文字还原
            self.btn_common.setText("只看常用9表")
            self.btn_common.setToolTip(
                "只显示日常最常用的 9 张报表"
                "（LK01/LK02/LK03/LK25/LK26/LK27/LK28/LK31/LK33）"
            )

        # 用"当前筛选框里的关键字"重新算一遍显示状态，
        # 这样常用模式与文字筛选可以叠加使用（互不冲突）。
        self._apply_filter(self.ed_filter.text())

        # 【关键】范围变了，右下角的"已选 N / 43"必须立刻重算。
        #   进入常用模式 → 分子只算 9 张内的勾选（如 9/43）；
        #   退回全表模式 → 分子恢复为全部勾选（如 43/43）。
        self._update_count()

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
        """
        刷新右下角的"已选 N / 总数"计数。

        【口径一：分母恒为 43】
            分母永远是模板里的正式报表总数（43 张），
            不随"只看常用9表"或关键字筛选而变化。
            这样用户随时能拿"已选数"与"模板总表数"对照。

        【口径二：分子只统计"当前选择范围内"的勾选】
            进入"只看常用9表"后，本次要生成的只有那 9 张，
            9 表以外的勾选与本次生成无关，因此不计入分子。
            所以刚进入常用模式时显示 9 / 43；
            在其中取消勾选 2 张，就变成 7 / 43。
        """
        n = len(self.selected_codes())
        total = len(self._sheets)
        if self._common_only:
            # 末尾补一句范围提示，避免用户误以为"43 张里只选了 9 张"是漏选
            self.lbl_count.setText(f"已选 {n} / {total} 张（常用表范围）")
        else:
            self.lbl_count.setText(f"已选 {n} / {total} 张")

    # ------------------------------------------------------------------
    def _scope_items(self) -> list[QListWidgetItem]:
        """
        返回"当前选择范围内"的列表项（保持列表顺序）。

        【什么是"选择范围"？】
            - 全表浏览模式：范围 = 全部 43 张；
            - 只看常用9表模式：范围 = 那 9 张常用表。
            9 表以外的项在常用模式下不属于选择范围，
            即使它们此前被勾选，也不会被生成。

        【新手必学：把重复逻辑抽成小函数（消除重复）】
            "某项算不算在当前范围内"这个判断，下面好几个方法都要用。
            与其到处复制 if self._common_only and code not in ...，
            不如集中写在这里。将来若要再加"只看核心5表"之类，
            只改这一处，其它方法自动跟着变。这叫"单一职责"。
        """
        items: list[QListWidgetItem] = []
        for i in range(self.list.count()):
            item = self.list.item(i)
            if self._common_only:
                code = str(item.data(Qt.UserRole) or "").upper()
                if code not in self._common_codes:
                    continue    # 不在常用 9 表内 → 不属于本次选择范围
            items.append(item)
        return items

    def selected_codes(self) -> list[str]:
        """
        返回本次要生成的报表代码（保持模板顺序）。

        【注意】这里遍历的是 _scope_items()，不是全部列表项。
            因此常用模式下，返回的只会落在 9 张常用表之内，
            与右下角"已选 N / 43"的分子完全一致——
            界面上显示几个，实际就生成几个，不会出现"看着 9 张、
            实际却生成了 43 张"的误导。

        【为什么"保持模板顺序"很重要？】
            报表之间有勾稽依赖，按模板(LK)顺序生成，前面的表先取数入缓存，
            后面的表就能复用，既快又符合业务逻辑。
            因为我们是按列表顺序遍历的，所以天然就是模板顺序。
        """
        codes = []
        for item in self._scope_items():
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