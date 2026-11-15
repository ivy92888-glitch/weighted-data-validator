# Weighted Data Validator

> 多源数据校验 + 加权评分 + 规则引擎 + 带颜色标记的检查报告生成器

A single-project Excel macro suite that turns messy, multi-column source data
into clean, trustworthy records. For anyone who has ever manually eyeballed
a spreadsheet full of free-text suggestions and thought "there must be a
better way".

---

## 这个工具解决什么问题

你有没有接过这样的数据表？每一行是一条待校正记录，旁边几个列
是别人写的**自由文本建议**——用关键词排优先级，
候选值和百分比参数藏在同一段字符串里，候选名称用 "/" 分隔了三四个备选，
链接里塞满了追踪参数，还有上百条业务规则要人工核对……

**Weighted Data Validator 把这件事一次性自动化：**

| 模块 | 它做什么 |
|---|---|
| **加权评分选最优值** | 同一个候选值被记录的**次数越多**、被记录的**时间越新**，得分越高。综合得分 = 频率 × 0.4 + 新旧程度 × 0.6 |
| **规则引擎扫描** | 按关键词匹配业务规则，命中则自动追加备注说明 |
| **建议文本结构化** | 从自由文本里自动提取出候选值和百分比参数 |
| **关键词优先级** | 自由文本里多个候选时，按关键词优先级排序 |
| **多备选取末值** | 用 "/" 分隔多个备选时，自动取最后一个 |
| **链接标准化** | 移除查询参数 → 提取实体标识符 → 统一域名后缀 |
| **描述精简** | 合并连续重复词、截断长分句、去掉共同前缀 |
| **颜色标记报告** | 建议修改的单元格自动上色，按建议类型分类输出到独立工作表 |

## 目录结构

```
weighted-data-validator/
├── src/
│   └── WeightedDataValidator.bas   ' 唯一代码文件 — VBA 宏
├── examples/
│   └── sample-data.xlsx            ' 虚构演示工作簿（含 3 张工作表）
├── docs/
│   └── Configuration.md            ' 常量说明 + 第二领域示例
├── README.md                       ' 本文档
├── LICENSE                         ' MIT
├── .gitignore
└── .gitattributes
```

## 快速上手

### 1. 环境

- 任何支持 VBA 的桌面表格软件（Microsoft Excel 365 / 2021 / 2019 等）
- 启用宏的工作簿格式（`.xlsm` 或 `.xlsb`）
- 首次使用需在文件 → 选项 → 信任中心 中启用 VBA 项目

### 2. 导入代码

1. 打开 VBA 编辑器（`Alt + F11`）
2. 在左侧项目树右键 → 导入文件
3. 选择 `src/WeightedDataValidator.bas`
4. 可选：设置对 `Microsoft Forms 2.0 Object Library` 的引用（工具 → 引用）

### 3. 准备你的工作簿

代码默认查找以下工作表，**列布局可在 `Configuration.md` 中调整**：

| 工作表 | 必需列 | 作用 |
|---|---|---|
| **Data**（主数据表） | A:ID / B:Raw Link / C:Category CN / D:Category EN / E:Suggested Name / F:Suggested Code / G:Suggested Rate / H:Notes | 待处理的原始数据 |
| **ReferenceCodes**（参考库） | A:Lookup Code / B:Code / C:Notes | 历史使用记录，供加权评分算法遍历 |
| **Rules**（规则配置） | A:Keyword / B:Annotation | 规则引擎的关键词 → 备注映射表 |

### 4. 运行

在 VBA 编辑器或宏列表里运行 `RunFullPipeline`。

```
运行流程：
  Initialize → ValidateWorkbook → RunConsistencyCheck
  → RunRuleScan → RunSmartSuggestion → SanitizeLinks
  → GenerateFeedbackReport
```

完成后，工作簿中会新增或刷新 `Feedback` 工作表，里面按分类写入
所有带建议的记录，并按优先级排序。

## 核心算法

### 加权评分（`CalculateBestCode`）

```
综合得分 = 频率得分 × WEIGHT_FREQUENCY + 新旧程度得分 × WEIGHT_RECENCY

其中：
  频率得分   = (该候选值出现次数 / 最大出现次数) × 100
  新旧程度得分 = ((最新行号 - 2) / (总行数 - 2)) × 100
```

**可调参数**：`WEIGHT_FREQUENCY` / `WEIGHT_RECENCY`（见 docs/Configuration.md）

这个比例为什么是 0.4:0.6？因为"**最近用过**"通常比"**用过几次**"更能说明
这个候选值的正确性——旧记录可能已经过时。如果你是库存管理场景，
可能想改成 0.5:0.5 或 0.6:0.4。

### 建议文本解析（`ExtractSuggestion`）

优先级顺序：

| 层级 | 关键词 | 匹配方式 |
|---|---|---|
| 主要（高） | `top` > `try` > `suggest` > `recommend` | 数组索引越小优先级越高 |
| 次要（低） | `may` > `let` > `allow` > `consider` | 主要关键词全未命中才检查 |

### 链接标准化（`SanitizeLink`）

```
原始链接：https://www.amazon.com/ref=xx/dp/B0CZXHF2X7?psc=1&pf_rd_...
    ↓ 移除查询参数
https://www.amazon.com/ref=xx/dp/B0CZXHF2X7
    ↓ 提取实体标识符（正则: [A-Z][0-9A-Z]{9}）
B0CZXHF2X7
    ↓ 重建统一域名后缀
https://amazon.example-domain/dp/B0CZXHF2X7
```

### 描述精简（`OptimizeDescription`）

连续重复词合并 → 长分句截断 → 去掉共同前缀。
把 "其他 > 其他 > 其他" 变成 "其他"。

## 配置要点

所有可调常量都在 `src/WeightedDataValidator.bas` 文件顶部，
同时也在 `docs/Configuration.md` 里有详细说明。**不要在 README 里抄数值**——
源码是唯一真源。

## 第二领域示例：电商 SKU 标准化

为了说明通用性，想象你把这个工具挪到一个完全不同的场景：

| 场景 A：产品属性校验 | 场景 B：地址归一化（完全不同的领域） |
|---|---|
| 属性值一致性检查 | **收件地址城市名归一化检查** |
| 受监管属性规则引擎 | **区域合规规则引擎**（邮编格式、州/省缩写等） |
| 历史属性值使用记录做参考库 | **历史地址记录做参考库** |
| amazon → amazon.example 域名统一 | **任意平台链接 → 产品页统一格式** |
| 带颜色标记的反馈工作表 | **同上——机制完全通用** |

同一套代码，同一套算法，只是换了工作表里的参考数据和规则。
这就是"clean-room 重写"的意义。

## 合规与免责

- 本项目由作者独立 clean-room 重写，不引用或拷贝任何前雇主的专有代码
  或业务文档。
- MIT 许可证仅覆盖作者本人贡献的实现，不构成对任何第三方权利的处分。
- 示例数据为完全虚构、中性的演示数据，与任何真实业务无关联。
- 与任何电子表格软件厂商无关联。各类电子表格产品名称仅作为常见运行环境
  名称提及，不代表背书。

## 文件清单

- `src/WeightedDataValidator.bas` — 唯一代码文件（VBA 宏）
- `examples/sample-data.xlsx` — 虚构演示工作簿
- `docs/Configuration.md` — 常量说明 + 第二领域示例
- `LICENSE` — MIT
- `.gitignore` — VCS 卫生
- `.gitattributes` — 行尾归一化

## 版本

- v1.0.0 — 2026-11-15 — 初始发布：加权评分、规则引擎、智能建议、链接清洗、描述精简、反馈报告生成

---

## Quick Summary (English)

Weighted Data Validator is a single-file Excel VBA toolkit that takes
messy spreadsheet records (free-text suggestions, mixed values,
partial matches, un-normalized URLs) and produces a clean, color-coded
**validation report** for every row. The core algorithm combines a
**frequency-recency weighted score** with a **rule engine** and a
**suggestion-text parser**. Designed as a generic framework — drop in
your own reference data and rules, then run. No proprietary dependencies.

## Maintenance Note

This README is intentionally **Chinese-primary**. The English section is
a concise summary only. All parameter values and constants are defined
in source code — **do not copy numbers into documentation**, as they
will drift out of sync. The source is the single source of truth.
