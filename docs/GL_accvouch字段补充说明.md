# GL_accvouch 字段补充说明（相对《GL_accvouch表字典》的修正与补充）

> **为什么需要这份文档？**
> 项目自带的 [`GL_accvouch表字典.xlsx`](../GL_accvouch表字典.xlsx) 是**手工整理的参考材料**，
> 现场已确认它"涵盖大部分字段，但可能不完整或有少许不够准确"。
> 若把字典当成唯一事实来源，就可能把"字典漏了"误判成"表里没有"，
> 进而写出错误的 SQL 或得出错误的排查结论。
>
> 本文件记录**已现场确认**的差异，以及**仍需核实**的可疑点。
> 规则：**以真实数据库为准，字典只作参考**；任何差异都记在这里，不散落在代码注释里。

---

## 一、已确认的字典缺口

### 1. `iyear` —— 会计年度（字典漏列，实际存在）

| 项目 | 内容 |
|---|---|
| 字段名 | `iyear` |
| 含义 | 会计年度（如 2026） |
| 字典状态 | **缺失**（xlsx 里查不到此字段） |
| 实际状态 | **存在**（现场确认） |
| 确认日期 | 2026-11-02 |
| 影响 | 无。现有 SQL 早已按 `WHERE iyear = ?` 过滤，写法正确 |

### 2. `iYPeriod` —— 年月整数（字典漏列，实际存在）

| 项目 | 内容 |
|---|---|
| 字段名 | `iYPeriod` |
| 含义 | 年月整数，如 202601 表示 2026 年 1 月 |
| 类型 | int |
| 字典状态 | **缺失**（xlsx 里查不到此字段） |
| 实际状态 | **存在**（现场查库确认） |
| 确认日期 | 2026-11-02 |
| 影响 | 无。这是本次新增 `gl_profit_amount_by_period()` 的**主要期间过滤字段** |

> **【重要纠正记录】**
> 我曾错误地断言"`iYPeriod` 是 `GL_accsum` 的字段，`GL_accvouch` 里没有"，
> 现场查库后确认 **`GL_accvouch` 也有 `iYPeriod`**。
>
> **教训**：字典没写 ≠ 字段不存在。凡涉及"某字段是否存在于某表"的判断，
> 都应以 `Database.list_columns()` 实查结果为准，不得凭"印象中它在别的表"下结论。

### 期间字段三者的关系（三者语义等价）

`GL_accvouch` 同时具备三种期间表达，可任选其一过滤：

| 字段 | 取值示例 | 特点 |
|---|---|---|
| `iyear` | 2026 | 只有年，需配 `iperiod` 用 |
| `iperiod` | 1 | 只有月，需配 `iyear` 用 |
| `iYPeriod` | 202601 | **年月合一**，单字段即可定位到"某年某月" |

**LK02 采用的写法**（`gl_profit_amount_by_period`）：

```sql
WHERE iyear = ?                      -- 冗余保险，排除跨年脏数据
  AND iYPeriod BETWEEN ? AND ?       -- 主过滤：[YYYY01, YYYY12]
  AND ccode IN (...)
```

**为什么用区间而非 `IN (202601,...,202612)`**：区间写更短、可读，
且 SQL Server 对 `BETWEEN` 能直接用索引。

> **【重要】这条缺口差点导致误判**
> 初次通读字典时，因查不到 `iyear`，一度怀疑现有 [`queries/u8.py`](../queries/u8.py)
> 的 [`gl_vouchers()`](../queries/u8.py:87)、
> [`gl_balance_by_codes()`](../queries/u8.py:160) 会报"列名无效"。
> 现场确认后排除。
>
> **教训**：字典不全 ≠ 字段不存在。凡"字典查不到但代码在用"的字段，
> 应标记为"待核实"而不是"有问题"；核实后把结论写进本文件。

---

## 二、仍需现场核实（尚未确认，不要凭猜写 SQL）

### 1. `iflag`（凭证标志，tinyint）取值含义未知

字典只写"凭证标志"，**未说明 0/1/2/… 各代表什么**
（未审核 / 已审核 / 已记账 / 已作废？）。

对财务报表的影响：若利润表要按"已记账凭证"口径取数，
就必须确知 `iflag` 的取值语义，否则 WHERE 条件无从写起。

**当前规避方式**：暂用 `cbook IS NOT NULL`（见下一条），不依赖 `iflag`。

### 2. `ibook`（记账标志）与 `cbook`（记账人）是否恒等

- `ibook`：tinyint，专门的"记账标志"位
- `cbook`：varchar，"记账人"姓名

现有 [`gl_vouchers()`](../queries/u8.py:87) 用 `cbook IS NOT NULL` 判断已记账。

**待确认**：是否存在"`ibook` 置位但 `cbook` 为空"（或反之）的记录？
若存在，用 `cbook` 判断会漏算或多算凭证。

**建议核实 SQL**（只读，可安全执行）：

```sql
SELECT ibook, COUNT(*) AS 行数,
       SUM(CASE WHEN cbook IS NULL THEN 1 ELSE 0 END) AS 记账人为空
FROM GL_accvouch
WHERE iyear = 2026 AND iperiod = 5
GROUP BY ibook
ORDER BY ibook;
```

若每组的"记账人为空"均为 0 或等于整组行数，说明两者语义一致，用哪个都行；
否则须以 `ibook` 为准（标志位比姓名更可靠）。

### 3. `cbook IS NOT NULL` 是否足以排除"未记账凭证"

U8 里 `cbook` 为空通常表示尚未记账，但**空字符串 `''` 与 `NULL` 是两回事**。
`cbook IS NOT NULL` 对 `''` 成立（`''` 不是 NULL），会把"记账人为空串"的记录算作已记账。

**建议核实 SQL**：

```sql
SELECT COUNT(*) AS 空串行数
FROM GL_accvouch
WHERE iyear = 2026 AND iperiod = 5 AND cbook = '';
```

若结果大于 0，则应把条件改为 `cbook IS NOT NULL AND cbook <> ''`。

---

## 三、字段速查（按 LK02 取数需要整理）

### 定位类

| 字段 | 含义 | 类型 |
|---|---|---|
| `i_id` | 自动编号（全表唯一） | int |
| `iyear` | **会计年度**（字典漏列） | int |
| `iYPeriod` | **年月整数**如 202601（字典漏列） | int |
| `iperiod` | 会计期间（月份） | tinyint |
| `csign` | 凭证类别字 | varchar |
| `isignseq` | 凭证类别排序号 | int |
| `ino_id` | 凭证编号 | smallint |
| `inid` | 行号（分录序号） | smallint |
| `dbill_date` | 制单日期 | datetime |

凭证唯一标识：`(iyear, csign, ino_id)`；分录唯一标识：再加 `inid`。

### 金额类（利润表核心）

| 字段 | 含义 | 类型 |
|---|---|---|
| `md` | **借方金额**（本位币） | money |
| `mc` | **贷方金额**（本位币） | money |
| `md_f` / `mc_f` | 外币借方 / 外币贷方金额 | money |
| `nfrat` | 汇率 | float |
| `nd_s` / `nc_s` | 数量借方 / 数量贷方 | float |
| `cexch_name` | 币种名称 | varchar |

> 财务报表面向人民币本位币，取 `md` / `mc` 即可。
> `_f` 后缀 = foreign（外币），`_s` 后缀 = 数量。

### 科目与辅助核算

| 字段 | 含义 |
|---|---|
| `ccode` | **科目编码**（最关键） |
| `ccode_equal` | **对方科目编码** |
| `cdept_id` | 部门编码 |
| `cperson_id` | 职员编码 |
| `ccus_id` | 客户编码 |
| `csup_id` | 供应商编码 |
| `citem_class` / `citem_id` | 项目大类编码 / 项目编码 |
| `cname` | 业务员 |

> **`ccode_equal` 的特殊价值**：
> 勾稽关系第 18 条要求「营业收入 = LK23 发票金额 − 记入营业外收入部分的开票收入」，
> 这类"必须剔除特定对方科目"的口径，只能靠 `ccode_equal` 精确识别，
> `GL_accsum`（科目汇总表）做不到。

### 状态类

| 字段 | 含义 | 备注 |
|---|---|---|
| `ibook` | 记账标志 | 取值语义待确认（见第二节） |
| `cbook` | 记账人 | 现有代码用它判断已记账 |
| `ccheck` | 审核人 | |
| `ccashier` | 出纳签字人 | |
| `cbill` | 制单人 | |
| `iflag` | 凭证标志 | 取值语义待确认 |
| `iflagbank` / `iflagPerson` | 银行账 / 往来账两清标志 | |
| `bdelete` | 是否核销 | |

### 摘要与业务

`cdigest`(摘要)、`cn_id`(票据号)、`dt_date`(票号发生日期)、
`csettle`(结算方式)、`idoc`(附单据数)

### 外部系统集成（财务报表面一般不用）

`coutaccset`、`coutbillsign`、`coutid`、`coutno_id`、`coutsign`、
`coutsysname`、`coutsysver`、`ioutperiod`、`ioutyear`、
`doutbilldate`、`doutdate`

### 自定义项（**不可硬编码解读**）

`cDefine1~cDefine14`(varchar)、`cDefine15`(int)、`cDefine16`(float)、
`cDefine4`/`cDefine6`(datetime)、`cDefine7`(float)、
`ctext1`/`ctext2`(凭证头自定义项)

> 同一编号在不同账套可能挂不同含义（U8 允许客户自定义），
> 取数时**不得**假定 `cDefine1` 一定是某个业务含义。

### 分录可编辑性（bit，与取数无关）

`bcodeedit`、`bCusSupInput`、`bDeptedit`、`bItemedit`、`bPCSedit`、
`bvalueedit`、`bvouchAddordele`、`bvouchedit`、`bvouchmoneyhold`

> 这些只控制"凭证录入界面能不能改"，不参与取数，不必理会。

---

## 四、取数路线选择（对 LK02 的影响）

利润表取"发生额"，有两条路：

| 路线 | 表 | 优点 | 缺点 |
|---|---|---|---|
| A | `GL_accsum` | 快；一行一科目一期间；已有 [`gl_balance_by_codes()`](../queries/u8.py:160) 可改造 | 无法按摘要/对方科目二次筛选 |
| B | `GL_accvouch` | 可加 `cdigest`、`ccode_equal`、辅助核算条件，精度高 | 数据量大，需 `SUM` 聚合，速度慢 |

**建议**：主体项目走 A（快），需要剔除/拆分的项目走 B（准）。
具体每行走哪条路，在需求逐条确认时按项目定，
并写进 [`reports/lk02_profit.py`](../reports/lk02_profit.py) 的对应配置注释里。

---

## 五、维护约定

1. **字典与真实库不一致时，以真实库为准**，并把差异补进本文件第一节。
2. 任何"仍待核实"的点，**不得凭猜写进 SQL**；先核实、再落码，核实结论回到本文件。
3. 核实用的 SQL 必须符合项目的只读约束
   （`SELECT` 开头，见 [`core/db.py`](../core/db.py:166) 的三道防线）。
4. 若后续拿到更完整的字典，请更新 [`GL_accvouch表字典.xlsx`](../GL_accvouch表字典.xlsx)，
   并同步删除本文件中已解决条目，避免两份文档长期并存互相矛盾。