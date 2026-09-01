# Amazon SKU运营数据分析与自动诊断系统

> 基于 Python + Pandas + Excel 的 Amazon SKU运营数据分析与自动诊断项目

---

## 一、项目简介

本项目是一个面向 Amazon 运营场景的数据分析与自动化项目。

项目模拟 Amazon 运营助理在日常工作中对 SKU 销售、流量及广告数据进行整理和分析的场景，通过 Python 和 Pandas 对 300 条 SKU 运营数据进行自动化处理。

项目主要实现：

- Amazon SKU 原始运营数据读取
- 数据质量检查
- 缺失值检查
- 重复数据检查
- CTR、CVR、CPC、ACOS、ROAS 等核心指标计算
- 高 ACOS SKU 自动筛选
- SKU 异常问题诊断
- 根据诊断结果生成运营建议
- 自动导出 Excel 运营分析报告

项目核心思路：

原始数据
↓
数据质量检查
↓
数据清洗
↓
核心指标计算
↓
异常 SKU 筛选
↓
SKU 问题诊断
↓
运营优化建议
↓
Excel 报告输出

---

## 二、项目背景

Amazon 运营人员在日常工作中需要处理大量 SKU 数据，包括商品流量、点击、订单、销售额、广告花费、库存等信息。

如果完全依靠人工逐个查看 SKU，容易出现：

1. 数据整理效率较低
2. 高广告花费 SKU 难以及时发现
3. 异常 SKU 筛选依赖人工
4. SKU 问题判断缺乏统一标准
5. 重复性的运营数据处理工作较多

因此，本项目尝试使用 Python 将部分重复性的数据分析工作自动化。

通过建立基础的运营指标计算和规则化诊断逻辑，将：

**数据 → 指标 → 异常 → 原因 → 运营动作**

形成一个基础的数据分析闭环。

---

## 三、项目定位

本项目主要用于：

- Amazon 运营助理岗位求职作品集
- Python 数据分析能力展示
- Amazon 基础运营指标学习
- SKU 数据分析流程实践
- 电商运营数据自动化分析练习

本项目并非真实 Amazon 店铺后台项目。

由于目前暂无实际 Amazon 店铺后台操作权限，项目中的 300 条 SKU 数据为根据 Amazon 常见运营数据字段构建的模拟数据。

项目重点在于展示：

- 对 Amazon 基础运营指标的理解
- 对 SKU 数据分析流程的理解
- 使用 Python / Pandas 进行数据处理的能力
- 将数据分析结果转化为运营判断的能力

---

## 四、技术栈

本项目主要使用：

- Python
- Pandas
- OpenPyXL
- CSV
- Excel

### Python

用于完成：

- 数据处理
- 指标计算
- 条件判断
- 自动化分析
- Excel 文件输出

### Pandas

主要用于：

- CSV 数据读取
- DataFrame 数据处理
- 缺失值检查
- 重复数据检查
- 指标计算
- 条件筛选
- 排序
- Excel 数据导出

### OpenPyXL

本项目使用 Pandas 的 `to_excel()` 方法导出 `.xlsx` 文件时，作为 Excel 写入引擎使用。

---

# 五、项目目录结构

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