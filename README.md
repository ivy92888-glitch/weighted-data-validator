# 你手动对过这样的表吗？Weighted Data Validator 让它一次性自动化

![](./examples/sample-data-preview.svg)

有一天领导给你一个 Excel，说"帮我校校这些数据"。
你打开一看——每一行旁边是几个**自由文本建议列**：

> **E 列（建议名称）：** "工具车 / 运输推车 / Garden Wagon"（用 `/` 分隔了 3 个备选）
> **F 列（建议编码）：** "先用7326909890"
> **G 列（百分比参数）：** 藏在同一段字符串里——"建议用7326909890-2.7%"
> **B 列（链接）：** `https://www.amazon.com/ref=xx/dp/B0CZXHF2X7?psc=1&pf_rd_w=abc123`
> **旁边还有一张** 历史使用记录表——同一个编码用过 15 种不同写法，哪个才是对的？
> **还有上百条业务规则**要你一条一条人工核对……

你叹了口气，开始从头翻起。

**这个工具就是为了干掉这类重复劳动而生的。**
它在 Excel 里加一段 VBA 宏，运行一次就能帮你：

1. 从历史记录里自动挑出**最值得相信的候选值**（不是瞎猜——是"用过次数多 + 最近一次出现时间近"的加权评分）
2. 从自由文本建议里**结构化提取**出编码、百分比、最终名称
3. 按关键词优先级**自动挑对那条建议**（先用 > 尝试用 > 建议用 > 建议）
4. 链接里塞的参数**清干净**，实体标识符提出来，域名统一
5. 所有命中业务规则的记录自动**上色标记**，输出到一个独立的检查报告表里

---

## 🏗️ 它是怎么做到的

### 核心算法：频率 × 0.4 + 新旧程度 × 0.6

想象数据库里"历史使用记录"工作表——同一个值 A 出现过 3 次，最新一次在第 150 行；值 B 出现过 2 次，最新一次在第 180 行。哪个更可信？

加权评分就是干这个的：

```
值 A 得分 = (3/3) × 100 × 0.4  +  (148/198) × 100 × 0.6  =  84.8 分
值 B 得分 = (2/3) × 100 × 0.4  +  (178/198) × 100 × 0.6  =  80.6 分
```

A 胜出。

为什么是 0.4:0.6？因为**最近用过**通常比**用过几次**更能说明这个值的正确性——旧记录可能已经过时。你随时可以调比例（比如 0.5:0.5），见 docs/Configuration.md。

### 模块一览（6 个）

| 模块 | 干什么 | 一句话总结 |
|---|---|---|
| ✅ **加权评分选最优值** | 频率 × 0.4 + 新旧程度 × 0.6，自动挑最可信的候选值 | "历史记录里谁最常用、最近最常用？我选它" |
| 🎯 **关键词优先级建议解析** | 自由文本里多个建议时，按关键词优先级挑对那条 | "先用 > 尝试用 > 建议用 > 建议" |
| 📦 **建议文本结构化提取** | 从 "先用7326909890" 里抽出编码 `7326909890`；从 "7326909890-2.7%" 里抽出百分比 `2.7%` | "把一坨自由文本拆成结构化数据" |
| 🔗 **实体标识符 + 链接清洗** | 移除所有查询参数，提取标识符（正则 `[A-Z][0-9A-Z]{9}`），重建统一域名后缀 | "https://amazon.com/ref=xx/dp/B0CZXHF2X7?psc=1 → https://amazon.example-domain/dp/B0CZXHF2X7" |
| 🛡️ **规则引擎扫描** | 从规则表加载关键词 → 备注映射，命中自动追加备注 | "在 Rules 工作表里加一行关键词和备注，就多了一条自动规则" |
| 📝 **颜色标记检查报告** | 按"仅 A 建议 / 仅 B 建议 / 两者都有"分类写入 Feedback 工作表，建议单元格自动上色 | "跑完给你一个带颜色的体检报告" |

---

## 🔥 谁能用？适用场景（至少 5 种不同行业）

### 场景 1 — 跨境电商产品属性校验（原始场景）
每个电商 SKU 旁边几张表给了"建议分类编码"和"建议名称"，格式乱七八糟。跑一遍这个工具，编码、名称、百分比全给你结构化好，还给每个编码打"历史可信度"分。

### 场景 2 — 物流地址校正
你的收货地址池里同一个城市被不同人写了 12 种写法（"北京市"、"北京"、"Beijing"、"BJ"、"北京朝阳"……）。拿这张表当 Data、历史收件记录当 ReferenceCodes，跑一遍 → 自动选出现最多最新的那一个作为归一化值。

### 场景 3 — SaaS 用户属性补齐
用户表里一些人填了"技术栈：React 或 Vue 或 Next"（多选建议），另一些列写了"29-31岁左右"（带百分比）。工具自动拆出主值（取最后一个备选 `PickLastName`），给属性值打历史可信度分，规则引擎可以加"React 生态项目"之类的业务规则。

### 场景 4 — 金融交易标签清洗
历史交易标签表被多个团队维护，同一个交易类型出现了 8 种写法（"信用"、"信用支付"、"信用卡"、"credit_card"……）。工具自动归并 + 按历史使用频率给推荐，链接清洗模块可以处理任何 URL 形式的标识符。

### 场景 5 — 制造业部件选型推荐
你的零部件数据库里，同一个功能部件有多个可选型号，每个型号被"用过的次数 + 最近一次使用时间"不同。用加权评分自动推荐一个优先型号，规则引擎可以加"腐蚀性环境禁止用 XX 型号"之类的业务规则。

### 场景 6 — 任意一张"建议列表"表
只要你的表格**满足这个条件**：

> 一张主表（待处理数据） + 一张参考库表（历史记录） + 可选一张规则表（关键词 → 备注）

那这套工具就能用。代码里每个模块都可以独立开关，按你具体场景选用。

---

## 🚀 快速上手（5 步）

### 1. 环境
任何支持 VBA 的桌面电子表格软件（Microsoft Excel 365 / 2021 / 2019 等）。文件另存为 `.xlsm` 格式。

### 2. 导入代码
打开 VBA 编辑器（`Alt + F11`）→ 左侧项目树右键 → 导入文件 → 选 `src/WeightedDataValidator.bas`。

### 3. 准备你的表
默认 3 张工作表（列布局可调，见 docs/Configuration.md）：

| 工作表 | 必需列 | 干什么 |
|---|---|---|
| **Data** | A:ID / B:Raw Link / C:Category CN / D:Category EN / E:Suggested Name / F:Suggested Code / G:Suggested Rate | 待处理的主数据 |
| **ReferenceCodes** | A:Lookup Code / B:Code / C:Notes | 历史使用记录（加权评分的数据源） |
| **Rules** | A:Keyword / B:Annotation | 规则引擎配置（关键词命中就追加备注） |

> 示例数据见 `examples/sample-data.xlsx`（虚构的园艺工具领域，完全中性）。

### 4. 参考库表（ReferenceCodes）是可选的 —— 怎么准备

加权评分算法需要一张**历史记录表**来计算"同一个候选值出现多少次、最近一次出现在第几行"。

**这张表不随代码发布**——它是你自己的业务数据，可能有几千到几万行历史记录。
你可以直接从已有系统导出一张，也可以手动建一张。三个可选的表结构模板
（从简单到丰富）见 `docs/Configuration.md → Reference Database Schema Templates`。

**一张最小可用的参考库表**（3 列就够）：

| A: Lookup Term | B: Candidate Value | C: Notes |
|---|---|---|
| Garden Cart | 7326909890 | historical usage 2024 |
| Garden Cart | 7326909890 | historical usage 2024 (repeat) |
| Paper Lantern | 4823908580 | historical usage 2025 |
| … | … | 新记录永远放到底部（行号越大 = 越新） |

> 💡 **行号顺序是核心机制。** 加权算法从下往上扫表，
> 行号越大的记录被视为"越新"。所以你的参考库表**新记录永远追加到最后**，
> 不要插到中间。

### 5. 运行
宏列表里找到 `RunFullPipeline`，点运行。

```
初始化 → 验证工作簿结构 → 一致性检查 → 规则扫描
→ 智能建议解析 → 链接清洗 → 生成颜色标记的 Feedback 表
```

### 5. 看结果
Feedback 工作表会自动出现，按建议类型分类，建议单元格自动上色。

---

## 🎨 颜色图例（Feedback 工作表）

| 颜色 | RGB | 含义 |
|---|---|---|
| 🟢 浅绿 | (226, 239, 218) | 工具建议你修改的单元格 |
| 🔵 灰绿 | (127, 150, 152) | 错误部分（描述精简时被判定为不对的部分） |
| 🟠 橙 | (198, 89, 17) | 正确部分（描述精简后留下的正确值） |
| 🟡 黄 | (255, 255, 0) | Feedback 工作表表头 |

---

## 🛠️ 可调常量（源码顶部）

所有可调参数都在 `src/WeightedDataValidator.bas` 文件顶部的 `Private Const` 和数组里。改完保存即可生效。**不要在 README 里抄数值——源码是唯一真源。**

完整说明 + 第二领域示例见 `docs/Configuration.md`。

---

## ⚖️ 合规与免责

- 本项目由作者独立 **clean-room 重写**。不引用或拷贝任何前雇主的专有代码或业务文档。
- MIT 许可证仅覆盖作者本人贡献的实现，不构成对任何第三方权利的处分。
- 示例数据为完全虚构、中性的演示数据，与任何真实业务无关联。
- 与任何电子表格软件厂商无关联。各类电子表格产品名称仅作为常见运行环境提及，不代表背书。

---

## 📦 文件清单

```
weighted-data-validator/
├── src/WeightedDataValidator.bas   ← 唯一代码文件（VBA 宏）
├── examples/sample-data.xlsx        ← 虚构演示工作簿
├── docs/Configuration.md            ← 常量说明 + 第二领域示例
├── README.md
├── LICENSE                           ← MIT
├── .gitignore
└── .gitattributes
```

---

## 📅 版本

- **v1.0.0** — 2026-11-15 — 初始发布

---

## Quick Summary (English)

Weighted Data Validator is a single-file Excel VBA toolkit for **intelligent
data validation** across messy spreadsheets. You give it a Data sheet, a
Reference sheet of historical records, and an optional Rules sheet — it
returns a **color-coded Feedback report** with the best candidate picked
by a **frequency-recency weighted score**.

Core algorithms:
- `CalculateBestCode` — 0.4 × frequency + 0.6 × recency
- `ExtractSuggestion` — keyword-priority parsing (top > try > suggest > recommend)
- `ExtractFirstNumber` / `ExtractPercent` — regex from free text
- `PickLastName` — split by separator, take last
- `SanitizeLink` — strip query params → extract entity ID → normalize domain suffix
- `CheckRuleMatch` — dictionary-based keyword → annotation

Designed as a **generic framework**. Drop it into product attribute
validation, address normalization, SKU consistency, transaction tag
cleaning — any domain where you have "free-text suggestions + a history
of good values + business rules".

---

## Maintenance Note

This README is intentionally **Chinese-primary**. The English section above
is a concise summary only. All parameter values and constants live in the
source code — **do not copy numbers into documentation**, as they drift
out of sync. Source is the single source of truth.
