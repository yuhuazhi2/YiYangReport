# 亿阳财务报表生成工具（YiYangReport）

一键把用友 **U8 V18.0** 后台（SQL Server 2016）的财务业务数据，按
[`2026版新修订企业财务报模板.xls`](2026版新修订企业财务报模板.xls) 的格式，
生成**一个包含所选表页**的 Excel 报表文件。

- 适用环境：Windows 10 / Windows Server 2016，局域网或服务器本机运行。
- 核心原则：**绝不修改模板文件**，所有结果另存到输出目录。
- 交互原则：**尽量减少多轮对话**——连接信息一次录入并记忆，选表页一次完成。

---

## 一、快速上手

### 1. 安装依赖

```bat
pip install -r requirements.txt
```

另外需要本机环境具备：

| 组件 | 说明 |
|------|------|
| Microsoft Excel | 生成 `.xls` 报表的主引擎（win32com 驱动，可完整保留格式/公式/合并单元格） |
| ODBC Driver 17 for SQL Server | 连接 SQL Server 2016 的驱动 |

### 2. 启动程序

```bat
python main.py
```

### 3. 首次使用

1. 弹出「数据库连接设置」对话框，填写服务器地址、认证方式、数据库（U8 年度库，如 `UFDATA_001_2026`）。
2. 点击「载入库列表」可自动列出服务器上的数据库，并优先定位 `UFDATA*` 年度库。
3. 点击「测试连接」确认成功，再点「确定」。
4. 进入主窗口后，设置会计期间、账套号、输出目录。
5. 点击「选择表页并生成…」勾选需要的报表，或「生成全部表页」。
6. 生成完成后可直接「打开结果文件」查看。

> 连接信息（含加密后的密码）保存在 `%APPDATA%\YiYangReport\user_config.json`，
> 下次启动会自动回填，无需重复输入。

---

## 二、目录结构

```
YiYangReport/
├── main.py                    程序入口
├── requirements.txt           依赖清单
├── 2026版新修订企业财务报模板.xls   报表模板（只读，绝不修改）
│
├── config/
│   └── settings.py            连接与全局配置、密码加密、持久化
├── core/                      核心层
│   ├── logging_setup.py       日志 + SQL 审计
│   ├── db.py                  SQL Server 访问（连接复用）
│   ├── template.py            模板解析（提取 LK 代码与标题）
│   ├── excel_engine.py        Excel 生成引擎（win32com）
│   └── context.py             取数上下文与共享缓存
├── reports/                   报表层（每张表页一个类）
│   ├── base.py                报表基类
│   ├── registry.py            报表注册表（代码 -> 类）
│   ├── stub.py                占位报表（未实现的表页）
│   ├── header.py              通用页眉填写
│   └── lk01_sample.py         LK01 示例实现（编写范例）
├── queries/
│   └── u8.py                  用友 U8 取数 SQL（集中管理）
├── verify/
│   └── reconciliation.py      勾稽关系核对（45 条公式）
├── app/
│   └── generator.py           生成编排器（顺序生表流程）
├── ui/                        界面层（PySide6）
│   ├── connection_dialog.py   数据库连接对话框
│   ├── report_select_dialog.py 表页选择对话框
│   └── main_window.py         主窗口
├── docs/                      设计与取数策略文档
└── output/                    生成的报表（默认输出目录）
```

---

## 三、工作原理

```
UI（主窗口）
   │  选中表页代码列表
   ▼
app/generator.py（编排）
   │  1. 解析模板，得到 43 张报表清单
   │  2. 建立单一 DB 连接 + 共享缓存上下文
   │  3. 打开 Excel，仅保留选中表页
   │  4. 按 LK 顺序逐表 fetch -> write -> verify
   │  5. 另存为一个多表页文件
   ▼
core/excel_engine.py（输出 .xls）
```

**顺序生表 + 共享缓存**：按 LK 代码顺序生成，符合报表间的勾稽依赖；
多张报表复用同一批 U8 数据，避免重复查询，提升大数据量下的效率。

---

## 四、模板表页构成

模板共 46 个工作表：

- **3 个说明类**：报表封面、报表目录及报送时间、主要勾稽关系（生成时自动保留）。
- **43 张正式报表**：LK01 ~ LK33（含 LK02-1、LK12-1 等子表），
  以**工作表名中的 `LK` 代码为唯一标识**，配合名称区分。

> 注：模板目录仅列到 YY31，缺少 LK32、LK33，程序以代码为准。

---

## 五、开发与扩展

### 新增一张报表的取数逻辑

1. 在 [`queries/u8.py`](queries/u8.py) 增加对应 SQL 函数（返回 `(sql, params)`）。
2. 在 `reports/` 下新建报表类，继承 [`reports/base.py`](reports/base.py:25) 的 `BaseReport`，
   用 `@register` 装饰并实现 `fetch()` / `write()`。
3. 在 [`reports/registry.py`](reports/registry.py:75) 的 `load_builtin_reports()` 中导入该模块。

未实现的表页会自动使用 [`reports/stub.py`](reports/stub.py:22) 的占位报表，
仅复制表页、不填数，保证整体流程可运行。

### 取数与核对

- 取数一律走 `ctx.fetch(key, sql, params)`，自动获得缓存与 SQL 审计。
- 报表可在 `verify()` 中登记关键指标，供 [`verify/reconciliation.py`](verify/reconciliation.py) 做勾稽核对。
- 所有 SQL 记录到 `logs/sql_audit.log`，主窗口提供「查看 SQL 审计」入口。
- 需人工比对后台数据时，可用 `Database.list_tables()` / `list_columns()`
  与 `queries.u8.sample_rows()` 取样核查。

> ⚠ **重要**：`queries/u8.py` 中的表名/字段名为 U8 通用命名，
> 实际部署时须以现场 `UFDATA` 年度库的真实结构核对调整。
> 详见 [`docs/数据流与取数策略.md`](docs/数据流与取数策略.md)。

---

## 六、常见问题

**Q：提示未检测到 Microsoft Excel？**
A：生成 `.xls` 依赖本机 Excel。请安装 Excel，并确保 `pip install pywin32`。

**Q：连接失败？**
A：检查 SQL Server 是否允许 TCP/IP 连接、防火墙端口（默认 1433）、
以及登录账户是否有权访问 U8 年度库。

**Q：运行时报 `ModuleNotFoundError: PySide6`？**
A：执行 `pip install PySide6`。

**Q：会不会改坏模板文件？**
A：不会。程序以只读方式打开模板，结果另存到输出目录。

---

## 七、打包为可执行程序（可选）

```bat
pip install pyinstaller
pyinstaller -w -n 财务报表生成工具 main.py
```

生成的 exe 在 `dist/` 目录。首次运行仍需配置数据库连接。