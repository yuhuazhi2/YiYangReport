# -*- coding: utf-8 -*-
"""
程序入口（整个工具的"起点"）
============================================================

【新手须知】Python 程序从哪里开始运行？
    当你执行 `python main.py` 时，Python 会从上到下读取这个文件。
    文件里通常只做两件事：
      1. 定义函数（把一堆操作"打包"起来，此时并不执行）；
      2. 在文件末尾用 `if __name__ == "__main__":` 判断"是不是被直接运行"，
         是的话才真正调用函数开始干活。

【本文件的职责】
    把"启动程序"这件事按顺序做完：
      第1步 初始化日志、目录等基础设施
      第2步 读取上次保存的配置（含数据库连接信息）
      第3步 检查本机是否具备 Excel 环境
      第4步 如果没有可用连接信息，弹出"连接设置"对话框让用户录入
      第5步 打开主窗口，把控制权交给 Qt 事件循环

【为什么要"尽量减少多轮对话"？】
    因为用户（财务人员）不希望每次都要重新输入数据库密码。
    所以：只有"没有配置"或"配置连不上"时，才弹连接对话框。
"""
from __future__ import annotations
# ↑ 这是一个"未来导入"声明。它让 Python 在旧版本中也能使用新版本的
#   类型标注写法（例如用 list[str] 代替 List[str]）。
#   对初学者来说：加上它没坏处，可以让类型提示更简洁，且不影响运行。


import sys
# ↑ sys 是 Python 内置模块，用于与"解释器本身"交互。
#   这里要用到 sys.argv（命令行参数）和 sys.exit（退出程序）。


def _setup_stdout() -> None:
    """
    让中文能正常打印，避免 Windows 命令行出现乱码。

    【新手须知】
        Windows 的命令行默认编码常常是 GBK，而我们代码里写的是 UTF-8 中文。
        如果二者不一致，打印中文就会出现"锟斤拷"这类乱码。
        把标准输出重新设置为 UTF-8，就能正常显示中文。

        函数名前的下划线 `_` 是约定：表示"这是本模块内部使用的辅助函数"，
        外部一般不需要调用它。
    """
    try:
        # reconfigure 把输出流的编码改成 utf-8
        # 加了 type: ignore 是因为部分 Python 版本的类型检查器不认识这个方法
        sys.stdout.reconfigure(encoding="utf-8")  # type: ignore[attr-defined]
        sys.stderr.reconfigure(encoding="utf-8")  # type: ignore[attr-defined]
    except Exception:
        # 即使设置失败也不能让程序崩溃：这里选择"安静地忽略"
        # （有些运行环境没有 reconfigure 方法，属于正常情况）
        pass


def main() -> int:
    """
    主函数：按顺序完成启动流程。
    返回值是"进程退出码"：0 表示正常结束。

    【新手须知：为什么要用函数包起来？】
        如果把所有代码都平铺在文件顶层，程序一读到这里就会执行，
        既不好复用，也不好测试，更难读懂。
        用函数包起来后，逻辑有了名字、有了边界，可读性大大提升。
    """
    _setup_stdout()

    # ------------------------------------------------------------------
    # 为什么把 import 放在函数里面？
    # 【新手疑问】通常 import 都写在文件最上方，这里为何放在函数内？
    #   这叫"延迟导入"（lazy import）。原因是：
    #   1. 有些库（如 PySide6 界面库）加载较慢或可能未安装，
    #      放在真正要用的时候再导入，可以避免"没用界面也要加载界面库"。
    #   2. 保证前面 _setup_stdout() 先把中文环境准备好。
    # ------------------------------------------------------------------
    from PySide6.QtWidgets import QApplication, QDialog, QMessageBox

    from config.settings import ensure_dirs, load_config, save_config
    from core.excel_engine import check_excel_available
    from core.logging_setup import get_logger
    from ui.connection_dialog import ConnectionDialog
    from ui.main_window import MainWindow

    # 取得一个"日志器"。之后可以用 logger.info(...) 把运行信息记到文件里。
    logger = get_logger("main")

    # 确保输出目录、日志目录等文件夹存在（不存在则创建）
    ensure_dirs()

    # ------------------------------------------------------------------
    # QApplication 是什么？
    #   Qt 程序必须先创建一个 QApplication 对象。
    #   可以把它理解为"整个图形界面的总管家"：它负责接收鼠标/键盘事件、
    #   管理窗口、运行事件循环。一个程序只能有一个 QApplication。
    #   sys.argv 是命令行参数（Qt 会读取其中一些标准参数）。
    # ------------------------------------------------------------------
    app = QApplication(sys.argv)
    app.setApplicationName("YiYangReport")       # 程序内部名（仅内部标识用）
    app.setApplicationDisplayName("U8报表生成器")  # 界面上显示的名字（对用户可见）

    # ------------------------------------------------------------------
    # 第2步：读取上次保存的配置
    #   load_config() 会去 %APPDATA%\YiYangReport\user_config.json 找配置；
    #   找不到就返回一份"默认配置"。这样无论第一次还是第 N 次运行都不会报错。
    # ------------------------------------------------------------------
    config = load_config()
    logger.info("配置已加载：模板=%s，数据库=%s",
                config.template_path, config.connection.database)

    # ------------------------------------------------------------------
    # 第3步：检查 Excel 环境
    #   本工具生成 .xls 依赖本机安装 Microsoft Excel（通过 win32com 驱动）。
    #   这里只是"提醒"，即使没有 Excel 也继续启动，
    #   让用户能进界面看到提示，而不是程序一闪而过。
    # ------------------------------------------------------------------
    excel_ok, excel_msg = check_excel_available()
    if not excel_ok:
        # QMessageBox.warning 会弹出一个"警告"对话框
        # 第一个参数 None 表示"没有父窗口"（此处主窗口还没创建）
        QMessageBox.warning(
            None, "环境提示",
            f"{excel_msg}\n\n生成 .xls 报表需要本机安装 Microsoft Excel 与 pywin32。",
        )

    # ------------------------------------------------------------------
    # 第4步：判断是否需要弹出"连接设置"对话框
    #
    # 【判断逻辑说明】
    #   我们"想尽量不打扰用户"，所以默认先假设配置是好的；
    #   只有以下两种情况才真正需要弹窗：
    #     (a) 配置里压根没有数据库信息（第一次使用）；
    #     (b) 有信息，但实际连不上（换了网络/改了口令等）。
    # ------------------------------------------------------------------
    need_connect = not (config.connection.database and config.connection.server)
    # ↑ not(...and...) 的含义：
    #   只有当"数据库名"和"服务器地址"都非空时，括号内才是 True，
    #   need_connect 才会是 False（即不需要弹窗）。

    if not need_connect:
        # 已有完整配置，就"安静地"验证一次连接是否真的可用
        # try/except 用于捕获可能出现的异常，避免程序崩溃
        try:
            from core.db import Database
            ok, _msg = Database(config.connection).test_connection()
            need_connect = not ok   # 连不上 → 需要弹窗
        except Exception as exc:  # noqa: BLE001
            # 出现任何异常都当作"需要重新配置"
            logger.warning("静默连接验证异常：%s", exc)
            need_connect = True

    if need_connect:
        # 创建并弹出连接对话框。
        # 【重要】调用 dlg.exec() 会"阻塞"在这里，直到用户关闭对话框。
        # 用户点"确定"返回 QDialog.Accepted，点"取消"返回 QDialog.Rejected。
        dlg = ConnectionDialog(config.connection)
        if dlg.exec() != QDialog.Accepted:
            # 用户取消了：直接退出程序（返回 0 表示正常退出，不是错误）
            logger.info("用户取消连接设置，程序退出。")
            return 0

        # 把对话框里用户填写的连接信息，写回全局配置，并保存到磁盘
        config.connection = dlg.result_config
        save_config(config)
        logger.info("连接配置已保存：%s / %s",
                    config.connection.server, config.connection.database)

    # ------------------------------------------------------------------
    # 第5步：创建并显示主窗口，进入事件循环
    #
    # 【什么是事件循环？】
    #   app.exec() 会让程序"停在这里等待用户操作"：
    #   用户点按钮、拖窗口、关闭程序……Qt 都会收到对应事件并调用我们的处理函数。
    #   直到所有窗口关闭，app.exec() 才会返回一个数字，程序随之结束。
    #   可以把它想象成一个"永不停止的服务员"，一直等待并响应客人的需求。
    # ------------------------------------------------------------------
    win = MainWindow(config)
    win.show()          # 把窗口显示出来（默认是隐藏的）
    return app.exec()   # 开始事件循环，把返回码作为整个程序的退出码


# ----------------------------------------------------------------------
# 【Python 必学知识点：if __name__ == "__main__"】
#
#   __name__ 是 Python 自动给每个模块设置的一个特殊变量：
#     - 当该文件被"直接运行"（python main.py）时，__name__ 的值是 "__main__"；
#     - 当该文件被"当作模块导入"（import main）时，__name__ 的值是 "main"。
#
#   所以我们用这个判断来保证："只有直接运行本文件时才启动程序"，
#   而被别人 import 时不会意外地把整个界面弹出来。这是一种良好习惯。
# ----------------------------------------------------------------------
if __name__ == "__main__":
    # sys.exit(n) 会立刻结束程序，并把 n 作为退出码交给操作系统。
    # 习惯上：0 表示成功，非 0 表示出错。
    sys.exit(main())