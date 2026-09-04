# Amazon SKU数据分析与自动诊断系统

> 基于 Python + Pandas 构建的 Amazon SKU 数据分析与基础诊断项目

[![Python](https://img.shields.io/badge/Python-3.x-blue)](https://www.python.org/)
[![Pandas](https://img.shields.io/badge/Pandas-Data%20Analysis-orange)](https://pandas.pydata.org/)
[![License](https://img.shields.io/badge/License-Learning%20Project-lightgrey)](#免责声明)

## 一、项目简介

本项目模拟日常工作中处理 SKU 销售、流量和广告数据的场景，使用 Python + Pandas 对 300 条 SKU 商品数据进行自动化处理，并将分析结果输出为 Excel 报告。

项目实现了从 **原始数据 → 数据质量检查 → 核心指标计算 → 异常 SKU 筛选 → SKU 问题诊断 → 建议 → Excel 报告** 的基础分析流程。

### 核心功能

- 自动生成 300 条 Amazon SKU 模拟数据
- 保存并重新读取 CSV 原始数据
- 检查字段、数据类型、缺失值及重复行
- 自动计算 CTR、CVR、CPC、ACOS、ROAS、广告订单占比
- 按 ACOS 筛选高广告成本 SKU
- 对 SKU 进行规则化问题诊断
- 根据诊断结果生成基础运营建议
- 将全部 SKU、高 ACOS SKU、SKU 诊断结果分别导出到 Excel
- 使用项目相对路径，避免依赖个人电脑上的固定绝对路径

> **数据说明：** 本项目中的商品、SKU、ASIN、流量、订单、销售额和广告数据均为模拟数据，不代表真实 Amazon 店铺后台数据。项目主要用于学习、求职作品展示和数据分析能力实践。

---

## 二、项目背景

Amazon 运营工作中需要持续关注 SKU 的流量、点击、订单、销售额、广告花费等数据。当 SKU 数量较多时，完全依靠人工查看 Excel 容易产生重复劳动，也不利于快速定位异常商品。

因此，本项目尝试把部分重复性的数据分析工作程序化，将：

**数据 → 指标 → 异常 → 原因 → 运营动作**

形成一个基础的数据分析闭环。

---

## 三、项目定位

本项目定位为 **跨境电商运营/开发产品求职作品 + Python 数据分析实践项目**，重点展示：

- 对 Amazon 基础指标的理解
- 对 SKU 数据分析流程的理解
- Python / Pandas 数据处理能力
- 基础的数据异常识别能力
- 将数据结果转化为运营判断的能力
- 使用代码减少重复性数据处理工作的思路

本项目不声称具备真实 Amazon Seller Central 店铺后台操作经验，也未接入 Amazon SP-API 或真实广告后台。

---

## 四、技术栈

| 技术 | 用途 |
|---|---|
| Python | 项目开发语言、数据处理及自动化分析 |
| Pandas | CSV 读取、DataFrame 处理、指标计算、筛选及 Excel 导出 |
| OpenPyXL | Pandas 导出 `.xlsx` 时使用的 Excel 写入引擎 |
| CSV | 原始数据存储 |
| Excel | 分析结果及诊断报告展示 |
| Git / GitHub | 项目版本管理及代码展示 |

---

## 五、项目目录结构

```text
Amazon-SKU-Operation-Analysis/
│
├── README.md
│
├── requirements.txt
│
├── amazon_operation_analysis.py
│
├── data/
│   └── amazon_300_product_operations_raw.csv
│
└── output/
    └── Amazon_SKU运营诊断报告.xlsx
```

### 文件说明

- `README.md`：项目说明、指标定义、运行方法及分析逻辑
- `requirements.txt`：Python 第三方依赖库
- `amazon_operation_analysis.py`：项目完整执行脚本
- `data/amazon_300_product_operations_raw.csv`：300 条 SKU 模拟原始数据
- `output/Amazon_SKU运营诊断报告.xlsx`：程序运行生成的 Excel 分析报告

---

## 六、数据字段说明

原始 CSV 共包含 **300 行、16 个字段**。

| 字段 | 中文含义 | 说明 |
|---|---|---|
| SKU | 商品库存单位 | 商品唯一库存编码 |
| ASIN | Amazon 标准识别号 | 模拟商品识别码 |
| Product_Name | 商品名称 | 模拟商品名称 |
| Category | 商品类别 | 商品所属类目 |
| Price | 商品售价 | 单件商品售价 |
| Rating | 商品评分 | 商品评分 |
| Reviews | 评论数量 | 商品累计评论数量 |
| Sessions | 商品访问次数 | 商品详情页访问次数 |
| Impressions | 曝光量 | 商品/广告获得的曝光次数 |
| Clicks | 点击次数 | 获得的点击次数 |
| Orders | 订单数量 | 商品产生的订单数量 |
| Ad_Orders | 广告订单数量 | 由广告带来的订单数量 |
| Sales | 销售额 | 商品产生的销售金额 |
| Ad_Spend | 广告花费 | 广告投放成本 |
| Inventory | 库存数量 | 当前模拟库存 |
| Returns | 退货数量 | 模拟退货数量 |

---

## 七、核心指标及计算公式

### 1. CTR — 点击率

```text
CTR = Clicks / Impressions
```

用于观察曝光后产生点击的能力。

例如：

```text
Impressions = 10,000
Clicks = 200

CTR = 200 / 10,000 = 2%
```

CTR 偏低时，可进一步关注主图、标题、关键词相关性及广告展示位置等因素。

### 2. CVR — 转化率

本项目采用：

```text
CVR = Orders / Sessions
```

用于观察商品访问后产生订单的能力。

例如：

```text
Sessions = 1,000
Orders = 80

CVR = 80 / 1,000 = 8%
```

CVR 偏低时，可进一步检查价格、Review、Listing 内容、产品卖点、优惠活动等因素。

### 3. CPC — 平均点击成本

```text
CPC = Ad_Spend / Clicks
```

用于观察平均获得一次点击需要支付的广告成本。

例如：

```text
Ad_Spend = $100
Clicks = 200

CPC = $100 / 200 = $0.50
```

CPC 较高时，可进一步检查关键词竞争程度、Bid、搜索词质量及广告结构。

### 4. ACOS — 广告销售成本占比

```text
ACOS = Ad_Spend / Sales
```

用于观察广告花费占广告相关销售额的比例。本项目使用总销售额进行模拟计算。

例如：

```text
Ad_Spend = $200
Sales = $1,000

ACOS = 200 / 1,000 = 20%
```

本项目将 `ACOS > 30%` 作为高 ACOS SKU 的基础筛选条件。

> 实际情况中，ACOS 是否合理不能只看固定百分比，还需要结合产品毛利率、目标利润、产品生命周期及广告策略判断。

### 5. ROAS — 广告投入产出比

```text
ROAS = Sales / Ad_Spend
```

用于观察每投入 1 美元广告费用带来的销售额。

例如：

```text
Sales = $1,000
Ad_Spend = $200

ROAS = 1,000 / 200 = 5
```

### 6. Ad_Order_Rate — 广告订单占比

```text
Ad_Order_Rate = Ad_Orders / Orders
```

用于观察总订单中广告订单所占的比例。

---

## 八、数据质量检查

程序读取 CSV 后，会进行基础数据质量检查：

### 缺失值

```python
df.isnull().sum()
```

检查每个字段是否存在缺失值。

### 重复行

```python
df.duplicated().sum()
```

检查是否存在完全重复的数据记录。

### 数据类型

```python
df.dtypes
```

确认字段的数据类型是否符合后续计算需求。

本项目当前提供的模拟数据经过检查后，字段缺失值为 0，完全重复行数量为 0。

---

## 九、指标自动计算

程序使用 Pandas 对 300 个 SKU 一次性计算核心指标：

```python
df["CTR"] = df["Clicks"] / df["Impressions"]
df["CVR"] = df["Orders"] / df["Sessions"]
df["CPC"] = df["Ad_Spend"] / df["Clicks"]
df["ACOS"] = df["Ad_Spend"] / df["Sales"]
df["ROAS"] = df["Sales"] / df["Ad_Spend"]
df["Ad_Order_Rate"] = df["Ad_Orders"] / df["Orders"]
```

相比人工逐个计算，这种方式可以降低重复性工作，并方便后续继续增加新的指标和判断规则。

---

## 十、异常 SKU 筛选

项目首先以 ACOS 作为基础筛选指标：

```python
bad_ad = df[df["ACOS"] > 0.30].copy()
```

即筛选 ACOS 大于 30% 的 SKU。

随后按照 ACOS 从高到低排序：

```python
bad_ad = bad_ad.sort_values(
    "ACOS",
    ascending=False
)
```

这样可以优先查看 ACOS 较高的 SKU。

> **重要：** “高 ACOS”不等于“广告一定烧钱”。例如广告花费只有 `$2`、销售额 `$1` 时，ACOS 为 200%，但实际广告花费金额很低。因此实际运营判断还需要结合 `Ad_Spend`、`Sales`、`Orders`、`CPC`、`CVR`、`ROAS` 等指标综合分析。

---

## 十一、SKU 自动诊断逻辑

项目使用规则判断的方式对 SKU 进行基础诊断。

当前规则优先级如下：

| 判断条件 | 基础诊断 |
|---|---|
| ACOS > 30% 且 CPC > 10 | 高 ACOS + 高 CPC：重点检查关键词竞价 |
| ACOS > 30% 且 CVR < 8% | 高 ACOS + 低 CVR：重点检查 Listing 转化 |
| ACOS > 30% 且 CTR < 1% | 高 ACOS + 低 CTR：重点检查主图和广告相关性 |
| ACOS > 30% | ACOS 偏高：进一步检查广告结构 |
| 其他情况 | 正常 |

代码中的判断是一个**学习项目中的规则化示例**，阈值并非 Amazon 官方统一标准，实际运营时需要根据产品类目、价格、毛利、历史数据和广告目标进行调整。

### 诊断示例

#### 高 ACOS + 高 CPC

```text
可能问题：点击成本较高

建议进一步检查：
- 关键词 Bid
- 高花费低转化搜索词
- 广告活动结构
- 关键词竞争程度
```

#### 高 ACOS + 低 CVR

```text
可能问题：有流量和点击，但购买转化不足

建议进一步检查：
- Listing
- 产品价格
- Review / Rating
- 主图及产品图片
- 五点描述
- Coupon / Promotion
```

#### 高 ACOS + 低 CTR

```text
可能问题：曝光后点击不足

建议进一步检查：
- 主图
- 标题
- 关键词相关性
- 广告投放词
```

---

## 十二、自动生成运营建议

程序根据 `Diagnosis` 字段进一步生成 `Action` 字段，将数据指标转化为基础运营动作。

例如：

```text
高 ACOS + 高 CPC
→ 检查高花费低转化关键词，降低无效词 Bid，优化关键词投放

高 ACOS + 低 CVR
→ 检查主图、价格、Review、优惠券、五点描述和 Listing 转化

高 ACOS + 低 CTR
→ 检查主图、标题、关键词相关性及广告投放词

普通高 ACOS
→ 进一步检查广告活动、关键词、搜索词及预算分配
```

这样可以把单纯的“数据表”进一步转换成“问题定位 + 基础运营动作”。

---

## 十三、Excel 报告

程序使用 `pd.ExcelWriter()` 将结果输出为一个 Excel 文件：

```text
output/Amazon_SKU运营诊断报告.xlsx
```

Excel 当前包含 3 个 Sheet：

### Sheet 1：全部SKU

包含全部 300 个 SKU 及计算后的运营指标和诊断结果。

### Sheet 2：高ACOS SKU

只保留 ACOS > 30% 的 SKU，并按照 ACOS 从高到低排序。

### Sheet 3：SKU诊断报告

展示异常 SKU 的核心指标、诊断结果以及对应的基础运营建议。

---

## 十四、项目运行环境

建议环境：

```text
Python 3.x
```

项目依赖通过 `requirements.txt` 统一管理：

```text
pandas
openpyxl
```

其中：

- `pandas`：用于数据读取、清洗、计算、筛选和分析
- `openpyxl`：用于支持 Pandas 导出 `.xlsx` Excel 文件
- `os`、`random`：属于 Python 标准库，无需通过 pip 单独安装

---

## 十五、项目运行步骤

### 1. 克隆项目

```bash
git clone https://github.com/Vve-tyy/Amazon-SKU-Operation-Analysis.git
```

进入项目目录：

```bash
cd Amazon-SKU-Operation-Analysis
```

### 2. 安装依赖

```bash
pip install -r requirements.txt
```

如果电脑存在多个 Python 环境，也可以使用：

```bash
python -m pip install -r requirements.txt
```

### 3. 运行程序

```bash
python amazon_operation_analysis.py
```

### 4. 查看结果

程序运行结束后：

- 原始 CSV 保存在 `data/` 目录
- Excel 分析报告保存在 `output/` 目录

最终报告：

```text
output/Amazon_SKU运营诊断报告.xlsx
```

---

## 十六、项目路径设计

项目没有把个人电脑上的绝对路径写死，而是通过 `__file__` 获取 Python 脚本所在目录，再使用 `os.path.join()` 拼接项目路径：

```python
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

data_dir = os.path.join(BASE_DIR, "data")
output_dir = os.path.join(BASE_DIR, "output")
```

随后分别得到 CSV 和 Excel 路径：

```python
csv_path = os.path.join(
    data_dir,
    "amazon_300_product_operations_raw.csv"
)

excel_path = os.path.join(
    output_dir,
    "Amazon_SKU运营诊断报告.xlsx"
)
```

这样项目移动到其他电脑后，只要保持项目目录结构不变，就不需要修改原始文件路径。

---

## 十七、项目运行流程

```text
生成300条模拟SKU数据
        ↓
保存原始CSV
        ↓
重新读取CSV
        ↓
数据质量检查
        ↓
计算CTR / CVR / CPC
        ↓
计算ACOS / ROAS / Ad_Order_Rate
        ↓
筛选ACOS > 30%的SKU
        ↓
按ACOS降序排序
        ↓
SKU自动诊断
        ↓
生成运营建议
        ↓
导出Excel
        ↓
生成3个分析Sheet
```

---

## 十八、项目成果

通过本项目，实现了一个基础的 Amazon SKU 数据分析与自动诊断流程：

- 构建 300 条模拟 SKU 原始数据
- 建立包含 16 个字段的数据结构
- 使用 Pandas 完成 CSV 数据读取
- 完成缺失值、重复行和数据类型检查
- 自动计算 CTR、CVR、CPC、ACOS、ROAS、Ad_Order_Rate
- 自动筛选高 ACOS SKU
- 使用多指标规则进行基础问题诊断
- 根据诊断结果生成运营建议
- 自动生成 Excel 分析报告
- 将项目代码、原始数据和分析结果进行 GitHub 版本管理

---

## 十九、项目局限与后续优化

### 当前局限

1. 数据为模拟数据，不是真实 Amazon Seller Central 数据；
2. 尚未接入 Amazon SP-API；
3. 尚未接入真实 Sponsored Products 广告数据；
4. 当前诊断规则属于基础规则，阈值需要根据业务场景调整；
5. ACOS 未结合产品实际毛利率进行盈亏平衡分析；
6. 尚未进行关键词、搜索词及 Campaign 层面的深度分析。

### 后续计划

#### 自动生成运营日报

进一步实现：

```text
每日导入运营数据
        ↓
Python自动分析
        ↓
识别异常SKU
        ↓
生成运营建议
        ↓
自动输出Excel日报
```

---

## 二十、面试展示重点

本项目主要希望展示以下能力：

### 数据处理能力

能够使用 Python / Pandas 对结构化运营数据进行读取、清洗、计算和筛选。

### Amazon 基础理解

理解 Sessions、Impressions、Clicks、Orders、Sales、Ad Spend、CTR、CVR、CPC、ACOS、ROAS 等基础指标之间的关系。

### 数据分析思维

不是只关注单个指标，而是尝试通过：

```text
CTR
CVR
CPC
ACOS
ROAS
```

组合判断潜在问题。

### 自动化思维

将重复性的：

```text
数据读取
→ 指标计算
→ 异常筛选
→ 问题分类
→ 运营建议
→ Excel输出
```

通过 Python 程序化处理，提高重复性数据分析工作的效率。

---

## 二十一、作者

**姓名：罗嘉乐**  
**专业：软件工程**  
**求职方向：跨境电商运营/开发产品助理**

---

## 二十二、免责声明

本项目仅用于个人学习、技术实践及求职作品展示。

项目中的 SKU、ASIN、商品名称、销售数据、广告数据等均为模拟数据，不代表任何真实 Amazon 店铺、品牌或广告账户数据。
