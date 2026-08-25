# WeeklyAnalysisGenerate 解读

基于 `src/weekly/WeeklyAnalysisGenerate.bas` v82（13,055 行 / 193 个过程）通读整理。

这是整个工作簿里最大的模块，也是唯一一个把「读 CSV」当成工程问题的模块。它把每日 Sophis
快照变成一张周度 Lombard 报表，中间经过一张 16 字段的可审计暂存表。

| | |
| --- | --- |
| 行数 | 13,055 |
| 过程数 | 193 |
| 暂存表字段 | 16 |
| 输出集中度表 | 6（3 维度 × 2 口径） |

---

## 整体形状

入口只有一个：`GenerateWeeklyAnalysis`。它读 `Home!WeeklyEndDate`，解析出比较日（上月）和
YTD 日（上年末），加载五份快照数据，然后按 `WeeklyAnalysisLayout.Layout` 里的坐标把六个区块
写到同一张 *Weekly Analysis* 表上。

但真正的重量不在报表区块，而在 `BuildRiskGranularitySection` —— 它一个人占了从第 11,397 行
往后的篇幅，加上它依赖的证书展开、实体名规范化和参照表维护，超过全模块的三分之二。

```mermaid
flowchart TB
    P["Positions CSV<br/><small>LoadWeeklyPositionData</small>"]
    A["Accounts CSV<br/><small>LoadWeeklyAccountData</small>"]
    C["Certificates .xlsx<br/><small>FindLatestCertificateFile</small>"]

    R["六个报表区块<br/><small>Overview · Breakdown · New<br/>Ended · Entered · Pie</small>"]

    EQ["Equity<br/><small>Security Name</small>"]
    BD["Bonds<br/><small>Ticker → Issuer</small>"]
    FD["Funds<br/><small>Prefix → Parent</small>"]
    CE["Certificates<br/><small>递归展开</small>"]

    NORM["实体名规范化与合并<br/><small>NormalizeEntityKey · BuildCanonicalEntityNameMap</small>"]

    STAGE["<b>Risk Exposure — 16 字段暂存表</b><br/><small>NDG · AssetClass · ExposureName · AllocationWeight<br/>AllocatedValue · Geography · Sector · ResolutionSource · AccountScope</small>"]

    FULL["Full（全组合）<br/><small>扫描时只收 DPM 行<br/>扫完并入 Excl. DPM</small>"]
    EXD["Excl. DPM<br/><small>扫描时只收 Non-DPM 行<br/>Asset Class = GP 即 DPM</small>"]

    T1["Name × 2"]
    T2["Geography × 2"]
    T3["Sector × 2"]

    P --> EQ & BD & FD
    A --> R
    P --> R
    C --> CE
    EQ & BD & FD & CE --> NORM
    NORM --> STAGE
    STAGE --> FULL & EXD
    EXD -.并入.-> FULL
    STAGE -. 可复用：跳过解析层，直接重新聚合 .-> FULL
    FULL --> T1 & T2 & T3
    EXD --> T1 & T2 & T3
```

六个报表区块走的是直路：读进来、按资产类别汇总、写出去。左边这条才是模块的主干——四条解析
支路把每一笔持仓变成一到多条「暴露」记录，汇进同一张暂存表，再从那张表分出两个口径、三个
维度。

报表区块用到的八个资产类别集中在 `CollateralCategories()` 一处定义（字典键 + 列标题），
Collateral Breakdown、Entered Collateral、各处合计和饼图切片颜色都从它派生，所以增删一个
类别只改这一个数组。

---

## 为什么这个模块的读取层和别处不一样

其他模块的 CSV 读取是 `Split(txt, ";")` 加 `Val()`。这里不是。三处细节说明作者是被真实
数据咬过的：

**数字解析不用 `CDbl`。** `WeeklyCsvDouble` 先看 `VarType`：已经是数值类型就直接 `CDbl`；
是字符串就自己动手——去掉 €、不换行空格、百分号，识别括号负数和尾随负号，然后**比较最后一个
逗号和最后一个点的位置**来判断哪个是小数点、哪个是千分位，最后才 `Val`。代码里那句注释写得
很直接：不要对 CSV 字符串用 `CDbl`，它是 locale-sensitive 的。换一台区域设置不同的机器，
同一份文件会解析出不同的数字——这个模块是唯一防住了这件事的。

**BOM 按字节值判断。** `CleanWeeklyCsvField` 剥引号、还原双写引号、去 `U+FEFF`，然后再检查
一次开头三个字符的码位是不是 239 / 187 / 191。为什么不直接写 UTF-8 BOM 字符？注释说明了：
这样 VBA 源码本身可以保持 ASCII-safe。二进制读入时 BOM 会以三个 ANSI 字符的形态存活下来，
只能按字节抓。

**列名可以变。** `FindWeeklyHeaderIndex` 接受一组候选表头名（`"Asset Type / Classification"`、
`"Asset Type"`、`"Classification"`），用 `NormalizeHeader` 抹掉空格、下划线、括号、斜杠后
比对；全都找不到就退到一个固定列号；再不行才按 `Required` 决定报错还是返回 −1。Issuer 和
Additional Comment 是 `Required:=False` 的——这两列不在也能跑。

> `WeeklySourceLines` 是先把 CRLF / CR 归一成 LF 再切分的。`DeltaCalculation.ReadAllLines`
> 后来补上的就是这套写法。

---

## 核心：Risk Exposure 暂存表

整个风险分析的枢纽是一张有名字的中间表，不是内存里的字典。`RiskStageField` 这个 Enum 定义了
16 个字段，每条记录代表**「某个 NDG 的某笔持仓，对某个实体贡献了多少暴露」**。

| 字段 | 作用 |
| --- | --- |
| `SourceRow` | 回溯到 Positions CSV 的原始行号 |
| `ProductISIN` / `SecurityName` / `OriginalAssetType` | 原始持仓身份，未经任何映射 |
| `AssetClass` | 风险口径的资产类别（不同于组合口径，见下） |
| `ExposureName` | 规范化后的最终实体名 |
| `ExposureType` | Equity issuer / Bond issuer / Fund parent company / Certificate underlying … |
| `AllocationWeight` | 这笔持仓分配给该实体的比例。非证书恒为 1 |
| `AllocatedValue` | 持仓市值 × 权重，聚合时实际相加的数 |
| `Geography` / `Sector` | 从 Companies 表解析出来的维度值 |
| `ResolutionSource` | **这个名字是怎么来的** |
| `AccountScope` | `DPM` 或 `Non-DPM`，建行时就定好 |

`ResolutionSource` 是整个设计里最值得称道的一个字段。它记的不是数据，是**数据的来历**：
`"Position: Security Name"`、`"BondIssuers: Issuer Ticker"`、`"UnmappedEquities: ISIN"`、
`"Fallback"`。集中度报表出了争议数字时，能直接在这一列上排序，看有多少暴露是靠 fallback
撑起来的。

### 复用机制

`ShouldRebuildRiskStageTables` 先验证表的 schema 和数据是否完好，完好就把
`RiskStageReuseDescription`（里面带着表上记录的 as-of 日期，以及它和当前请求日期是否一致）
交给 `ShouldOverwriteExistingSheets` 去问用户。选择复用时，`BuildRiskGranularitySection`
直接 `GoTo` 到聚合标签——整个证书文件加载、参照表更新、实体解析全部跳过，只从表里重新聚合。

> **操作上的一个后果**：这个询问对话框是在 `ScreenUpdating = False` 的状态下弹出来的，
> 而且发生在报表跑到一半（`BuildRiskGranularitySection` 排在 `BuildNewLoansSection` 之前）。
> 点了「Weekly Analysis」之后界面会先静默一段，然后突然弹框问是否重建。

---

## 证书递归展开

模块里最复杂的算法是 `ExpandCertificateRic`。输入一个 RIC，输出一个「成分 + 权重」的集合，
权重之和为 1。结构性证书的底层可能是另一个篮子，篮子里还可能有篮子，所以它是递归的。

1. **四道熔断。** 空 RIC、深度超过 `CERTIFICATE_MAX_BASKET_DEPTH`（12）、`RecursionPath`
   里已存在（成环）、RIC 在映射表里找不到——四种情况都产出一个带 `__MISSING_RIC__|` 前缀、
   权重 1、标记为 Temporary 的占位成分。不抛错，不静默丢弃。
2. **不是篮子就到底了。** `IsBasketUnderlyingName` 判否，直接产出这个实体，权重 1。
3. **是篮子且有子 RIC 列表。** 把子列表按逗号、分号、竖线、各种换行切开，去重，按 `1 / n`
   等权递归展开。
4. **关键的回退。** 如果所有子 RIC 展开完，`CertificateComponentsContainResolvedEntity`
   发现**一个真实体都没有**，就把整个结果丢掉重来，改用篮子的描述性文本解析。代码注释解释了
   原因：过期的成分 RIC 会从新版 InstrumentList 里消失，这时候那段人类可读的文本反而更靠得住。
   这一条不是设计出来的，是从事故里长出来的。
5. **文本路径又是三级。** 先在文本里直接找已知 RIC；找不到就抽出第一个 ISIN 查 ISIN→RIC；
   再不行就 `CleanBasketComponentName` 清洗出名字查 名字→RIC；全都不行才把清洗后的名字直接
   当成实体。

一个容易被忽略但很重要的细节：函数返回前会执行 `RecursionPath.Remove Ric`。这是**正确的深度
优先回溯**，不是一个简单的 visited 集合。同一个标的出现在篮子的不同分支里是完全合法的（会被
各自的权重分别计入），只有真正的自我引用环才会被熔断。用 visited 集合的话，第二次出现就会被
误判成环。

---

## 实体名规范化

Sophis 的 Security Name、Bloomberg 的 Issuer Name、证书篮子里的文本描述、Companies 表里的
手工名——同一家公司在四个来源里有四种写法。`NormalizeEntityKey` 是一条流水线：

- 去掉尾部的 `" Class A"` 这类股份类别。判定器 `IsEntityShareClassSuffix` 只承认 1–2 个字符
  的类别代码（可以后接 `SHARE` / `ORDINARYSHARES` 等词），注释明说是为了不把 `Inc` 误认成
  类别代码。
- 去掉 Bloomberg 风格的 `/The`、`, THE` 和前导 `The `。
- `&` → `AND`，去变音符号，只保留字母和数字。
- 最后归一法律后缀。

在这之上还有一层模糊合并 `IsLikelyEntityPrefixMatch`，两条规则：一个键是另一个键去掉
`GMBH` / `CORP` / `INC` / `SPA` 等后缀的结果（无条件合并）；或者短键是长键的前缀，且短键至少
`ENTITY_PREFIX_MIN_LENGTH`（12）个字符、长度比不低于 `ENTITY_PREFIX_MIN_RATIO`（0.65）。

> **正确的保守选择**：`FindEntityPrefixCanonicalName` 遍历所有已知键找前缀匹配。如果匹配到
> **两个不同的** canonical name，它判定为 ambiguous，然后**放弃合并**——宁可让同一家公司暂时
> 分成两行，也不把两家公司合成一行。对一张要拿去看集中度的表来说，这个取舍方向是对的。

---

## 三个分类函数，各管各的

这是最容易被后来的人改错的地方。同一笔持仓在「组合构成」和「风险集中度」两个口径下**可以属于
不同类别**，而且这个差异是被显式声明的，不是意外。

| 函数 | 管什么 | 依据 |
| --- | --- | --- |
| `GetAssetClass`（AssetMapping） | Collateral Breakdown、饼图、Entered Collateral | Asset Type 的 `Like` 前缀匹配 |
| `ResolveTopTenAssetClass` | **只**用于 Top 10 集中度 | 基类是 GP 时，按 `Additional Comment` 的前缀再拆成 Equity / Bonds / Funds |
| `IsRiskRelevantCertificateAssetType` | 证书是否进入风险分析 | 原始 Asset Type 归一后必须等于 `CERTIFICATESOTHER` |

`ResolveTopTenAssetClass` 上方有一条注释，明令禁止把它用在核心分类里；
`IsRiskRelevantCertificateAssetType` 上方那条则声明自己是 "authoritative risk-analysis
filter"，并说明它**故意**基于原始 Sophis Asset Type 而非 `GetAssetClass`。这两条注释是这个
模块里最重要的两条。

### DPM 到底是什么

`RiskPositionCacheIsDPM` 的定义就一行：`GetAssetClass(AssetType) = "GP"`，也就是 Segregated
Account（全权委托）。于是出现了一个效果：`ResolveTopTenAssetClass` 会把一笔 GP 里的股票重新
归到 Equity，所以它**在 Full 表的 Equity 子表里出现**；而在 Excl. DPM 表里，它整个消失。
左右两张表的差额，正好就是全委账户穿透后的贡献——这正是这两张表并排放在一起要回答的问题。

---

## 聚合与开关

### 一行只更新一套字典

`AggregateUnifiedRiskStageData` 里有个乍看反直觉的写法：扫描过程中，`Full*` 数组**只**接收
DPM 行，`ExDPM*` 数组只接收 Non-DPM 行。扫完之后再 `MergeRiskAggregateSet` 把 ExDPM 并进
Full。这样每一行只做一套字典更新而不是两套——版本注释里的 v81 就是这件事。扫描末尾还有个对账
断言：两边行数之和对不上暂存表总行数就直接抛错。

### 注释即配置

`BuildRiskSubtableVisibility` 是一张字典，键是 `维度|资产类`。注释掉一行，那个子表就从两个
视图里消失，**同时**从对应的 Overall 分母和 Top 10 计算里剔除——这个联动是 v53 明确加进去的，
为的是避免子表和 Overall 口径打架。当前状态：

| 维度 | Equity | Corp Bonds | Sov Bonds | Funds | Certificates | Overall |
| --- | --- | --- | --- | --- | --- | --- |
| Issuer | 开 | 开 | 开 | 开 | 开 | 关 |
| Country | 开 | 开 | 关 | 关 | 开 | 关 |
| Sector | 开 | 开 | 关 | 关 | 开 | 关 |

### Same as left

`ExposureGroupsMatch` 会比较 Full 和 Excl. DPM 两侧的 Certificates 结果。完全一致时右表那一段
用一行 "Same as left" 代替整块重复内容（v59）。证书通常不进全委账户，所以这一段大多数时候确实
是重复的。

---

## 参照表是自我维护的

`UpdateRiskReferenceDatabases` 在每次重建之前扫一遍持仓，往三张手工表里补空缺。规则是
append-only，已有的人工填写不会被覆盖——**只有一个例外**：

- **Equity Names**：只补「有 ISIN 但 Security Name 为空」的持仓，写进 `UnmappedEquities`
  等人工填名字。
- **Bond Issuers**：按 ticker 一行。各字段独立累积，所以同一 ticker 下的另一个 ISIN 可以补上
  缺失的 issuer 或分类，不用多扫一遍表。`Issuer Name` 是**唯一允许覆盖已有值**的字段，注释
  给了理由：Sophis 之后可能给出更正过的名字。
- **Funds**：补 Reference ISIN，并给 Prefix / Company Name 写 XLOOKUP 公式。
  `SetFundLookupFormula` 先试 `Formula2`（逗号分隔），失败再退到 `FormulaLocal`（分号分隔）
  ——处理的是 Excel 区域设置差异。

---

## 值得盯的几个地方

都不是 bug，是这个规模的模块里值得留意的结构性风险。

**性能：三处 O(n²)。** `WriteTopExposureGroup` 和 `WriteAssetTypeMapping` 都用冒泡排序。前者的
n 是某个资产类别下的唯一实体数——为了取前 10 名而把全表冒泡一遍。更重的是
`FindEntityPrefixCanonicalName`：每遇到一个新实体键就遍历一次全部已有键，整体 O(n²) 的字符串
比较。版本注释里 v64 到 v76 一直在做性能优化（缓存重复的证书展开、避免重复合并地域候选、缓存
实体排序用的比较属性），这三处是剩下的量级项。

**Notes 框的高度由饼图决定。** `BuildNotes` 的 LastRow 是
`Layout.PieRow + Layout.PieHeightRows - 1`。改饼图的高度会连带改掉 Notes 框的高度。两者在
Layout 里是独立字段，这层耦合只存在于 `BuildNotes` 这一行里。

**Additional Comment 的前缀判定有顺序依赖。** `ResolveTopTenAssetClass` 按 `EQUITY`(6) →
`FIXEDINCOME`(11) → `FUND`(4) 的顺序做前缀匹配。`"EQUITY FUND"` 会被判成 Equity，
`"FUND EQUITY"` 会被判成 Funds。这取决于 comment 实际的书写约定——如果那边的写法有变，Top 10
里 GP 的拆分方式会跟着变，而且没有任何提示。
