# 测试说明

## 一、怎么运行

在项目根目录执行：

```bat
python -m unittest discover -s tests -v
```

若已安装 pytest，也可以用：

```bat
python -m pytest tests -v
```

两者都能识别本目录下的测试文件，因为用的是标准 `unittest` 写法。

**运行条件**：只需要 Python 和项目依赖（xlrd）。**不需要** SQL Server，也**不需要** Excel。

## 二、为什么不需要数据库和 Excel

报表生成的真实链路是：

```
fetch(ctx)   → 查 SQL Server
write(hook)  → 操作真实 Excel
```

测试里用两个"假对象"顶替（见 [`fakes.py`](fakes.py)）：

| 假对象 | 顶替 | 行为 |
|---|---|---|
| `FakeCtx` | `ReportContext` | `fetch()` 直接返回预设的行，不连数据库 |
| `FakeHook` | `_SheetHook` | `set()` 把值记进字典，不启动 Excel |

所以测试跑得极快（全部 12 个用例约 0.06 秒），且可在任何机器上运行。

## 三、测了什么

### 1. 业务逻辑（`TestRowValues`）

| 用例 | 验证内容 |
|---|---|
| `test_all_rows_may` | 5 月：每行期末取本期 `me`、年初取 1 月 `mb` |
| `test_january_same_row_two_values` | 1 月：同一行同时提供期末与年初，互不干扰 |
| `test_money_fund_excludes_1012` | 货币资金排除 1012（1012 归交易性金融资产） |
| `test_no_data_writes_nothing` | 无记录时**不写 0**，保持空白 |
| `test_partial_data` | 只有本期、无 1 月记录时，只写期末格 |
| `test_fetch_failure_not_crash` | 取数抛异常时不崩溃 |
| `test_codes_deduplicated_in_query` | 所有行的科目合并成**一次**查询 |

### 2. 会计期间（`TestPeriodCell`）

| 用例 | 验证内容 |
|---|---|
| `test_period_written_as_text` | C4 写入"2026年5月"，且必须用 `set_text`（用 `set` 会被 Excel 识别成日期 46143） |

### 3. 结构校验（`TestConfigConsistency`）—— 长期护栏

这组用例不是"跑一遍看对不对"，而是**每次改代码都会自动检查**的护栏：

| 用例 | 防什么问题 |
|---|---|
| `test_cell_map_matches_specs` | 新增行时忘了登记 `cell_map` → 该行不进勾稽核对，核对"通过"是假象 |
| `test_no_duplicate_cells` | 两行占用同一格子 → 后写的覆盖先写的 |
| `test_codes_are_strings` | 科目编码写成数字 → 可能匹配不上 |
| `test_cells_in_template_have_labels` | 地址写错行（如 C21 写成 C12）→ 该行模板 A 列是空的，立刻暴露 |

**最后一条尤其重要**：它直接读模板 `.xls`，检查配置里每个格子所在行的 A 列有没有项目名。若模板不存在则自动跳过，不影响其他测试。

## 四、不测什么

**真实 Excel 的写入效果不在这里测**，包括：

- 数字格式是否正确
- 日期是否被误判成序列号（C4 踩过的坑）
- 合并单元格是否被破坏
- 行高列宽是否正常

原因：这些问题只有真 Excel 才会暴露，`FakeHook` 模拟不出来。这类验证必须手工做——生成一次真实文件，打开看。

判断标准：**逻辑对不对看测试，样子对不对用眼睛。**

## 五、新增一张报表时怎么做

以 LK02 为例：

1. 新建 `reports/lk02_profit.py` 实现报表类
2. 新建 `tests/test_lk02_logic.py`，照 [`test_lk01_logic.py`](test_lk01_logic.py) 的结构写：
   - 业务逻辑用例（取数、汇总、缺数据）
   - 结构校验用例（`cell_map` 一致性、模板地址有效性）
3. 跑一遍测试，确认通过

这样后续改动共用框架（如 [`queries/u8.py`](../queries/u8.py) 的 `gl_balance_by_codes`）时，跑一遍测试就能知道**有没有碰坏已完成的报表**，不必逐张重新生成核对。