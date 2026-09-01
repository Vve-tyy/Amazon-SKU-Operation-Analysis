# ============================================================
# Amazon SKU运营数据分析与自动诊断系统
# 功能：
# 1. 生成300条Amazon SKU模拟运营数据
# 2. 保存为CSV
# 3. 读取CSV
# 4. 数据质量检查
# 5. 计算CTR、CVR、CPC、ACOS、ROAS
# 6. 筛选高ACOS SKU
# 7. 自动诊断SKU问题
# 8. 自动生成运营建议
# 9. 导出最终Excel报告
# ============================================================


# ============================================================
# 第一部分：导入Python库
# ============================================================

import pandas as pd
import random
import os

# ============================================================
# 第二部分：设置文件路径
# ============================================================

# Windows路径前面加 r，避免反斜杠被Python特殊处理

# folder_path = r"D:\python练习\数据practice"

# # 原始CSV文件
# csv_path = folder_path + r"\amazon_300_product_operations_raw.csv"

# # 最终Excel报告
# excel_path = folder_path + r"\Amazon_SKU运营诊断报告.xlsx"

#项目相对路径，便于找寻
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

data_dir = os.path.join(BASE_DIR, "data")
output_dir = os.path.join(BASE_DIR, "output")

csv_path = os.path.join(
    data_dir,
    "amazon_300_product_operations_raw.csv"
)

excel_path = os.path.join(
    output_dir,
    "Amazon_SKU运营诊断报告.xlsx"
)

# ============================================================
# 第三部分：生成300条Amazon SKU模拟数据
# ============================================================

# 商品分类
categories = [
    "Electronics",
    "Home & Kitchen",
    "Sports & Outdoors",
    "Beauty",
    "Tools & Home Improvement",
    "Pet Supplies",
    "Toys & Games",
    "Office Products"
]

# 商品名称
product_names = [
    "Charging Cable Set",
    "Wireless Mouse",
    "Storage Hooks Set",
    "Camping Lantern Basic",
    "Silicone Spatula Set",
    "Travel Cosmetic Bag Premium",
    "Phone Stand",
    "LED Desk Lamp",
    "Kitchen Storage Rack",
    "Pet Grooming Brush",
    "Yoga Mat",
    "Water Bottle",
    "Bluetooth Speaker",
    "Laptop Stand",
    "Makeup Organizer",
    "USB Hub",
    "Travel Backpack",
    "Resistance Bands",
    "Desk Organizer",
    "Cleaning Brush Set"
]


# 用列表保存300条商品数据
data = []


# 循环生成300个SKU
for i in range(1, 301):

    # SKU编号
    sku = f"SKU-{i:04d}"

    # 模拟ASIN
    asin = "B" + "".join(
        random.choices("0123456789", k=9)
    )

    # 随机商品名称
    product_name = random.choice(product_names)

    # 随机类目
    category = random.choice(categories)

    # 商品价格
    price = round(random.uniform(8, 80), 2)

    # Rating
    rating = round(random.uniform(3.0, 5.0), 1)

    # 评论数量
    reviews = random.randint(20, 10000)

    # 会话数
    sessions = random.randint(300, 5000)

    # 曝光量
    impressions = random.randint(3000, 15000)

    # 点击量
    clicks = random.randint(20, 500)

    # 订单量
    orders = random.randint(10, 600)

    # 广告订单
    ad_orders = random.randint(5, min(orders, 300))

    # 销售额
    sales = round(price * orders, 2)

    # 广告花费
    ad_spend = round(
        random.uniform(50, 5000),
        2
    )

    # 库存
    inventory = random.randint(50, 2000)

    # 退货数量
    returns = random.randint(0, 30)


    # 把一条商品数据放进列表
    data.append([
        sku,
        asin,
        product_name,
        category,
        price,
        rating,
        reviews,
        sessions,
        impressions,
        clicks,
        orders,
        ad_orders,
        sales,
        ad_spend,
        inventory,
        returns
    ])


# ============================================================
# 第四部分：把数据转换成DataFrame
# ============================================================

columns = [
    "SKU",
    "ASIN",
    "Product_Name",
    "Category",
    "Price",
    "Rating",
    "Reviews",
    "Sessions",
    "Impressions",
    "Clicks",
    "Orders",
    "Ad_Orders",
    "Sales",
    "Ad_Spend",
    "Inventory",
    "Returns"
]

df = pd.DataFrame(
    data,
    columns=columns
)


# ============================================================
# 第五部分：保存原始CSV
# ============================================================

df.to_csv(
    csv_path,
    index=False,
    encoding="utf-8-sig"
)

print("=" * 60)
print("第一阶段完成：300条SKU数据已经生成")
print("=" * 60)

print("CSV文件位置：")
print(csv_path)


# ============================================================
# 第六部分：重新读取CSV
# ============================================================

df = pd.read_csv(csv_path)

print("\n" + "=" * 60)
print("第二阶段：读取CSV数据")
print("=" * 60)

print("数据形状：")
print(df.shape)

print("\n前5行数据：")
print(df.head())


# ============================================================
# 第七部分：数据质量检查
# ============================================================

print("\n" + "=" * 60)
print("第三阶段：数据质量检查")
print("=" * 60)


# 1. 查看字段名称

print("\n字段名称：")
print(df.columns)


# 2. 查看缺失值

print("\n缺失值检查：")
print(df.isnull().sum())


# 3. 查看数据类型

print("\n数据类型：")
print(df.dtypes)


# 4. 检查重复行

duplicate_count = df.duplicated().sum()

print("\n重复行数量：")
print(duplicate_count)


# ============================================================
# 第八部分：计算Amazon核心运营指标
# ============================================================

print("\n" + "=" * 60)
print("第四阶段：计算Amazon核心运营指标")
print("=" * 60)


# ------------------------------------------------------------
# CTR 点击率
# CTR = Clicks / Impressions
# ------------------------------------------------------------

df["CTR"] = df["Clicks"] / df["Impressions"]


# ------------------------------------------------------------
# CVR 转化率
# CVR = Orders / Sessions
# ------------------------------------------------------------

df["CVR"] = df["Orders"] / df["Sessions"]


# ------------------------------------------------------------
# CPC 平均每次点击成本
# CPC = Ad_Spend / Clicks
# ------------------------------------------------------------

df["CPC"] = df["Ad_Spend"] / df["Clicks"]


# ------------------------------------------------------------
# ACOS 广告销售成本占比
# ACOS = Ad_Spend / Sales
# ------------------------------------------------------------

df["ACOS"] = df["Ad_Spend"] / df["Sales"]


# ------------------------------------------------------------
# ROAS 广告投入产出比
# ROAS = Sales / Ad_Spend
# ------------------------------------------------------------

df["ROAS"] = df["Sales"] / df["Ad_Spend"]


# ------------------------------------------------------------
# 广告订单占比
# Ad_Order_Rate = Ad_Orders / Orders
# ------------------------------------------------------------

df["Ad_Order_Rate"] = df["Ad_Orders"] / df["Orders"]


# ============================================================
# 第九部分：查看计算结果
# ============================================================

print("\n核心运营指标：")

print(
    df[
        [
            "SKU",
            "CTR",
            "CVR",
            "CPC",
            "ACOS",
            "ROAS",
            "Ad_Order_Rate"
        ]
    ].head(10)
)


# ============================================================
# 第十部分：格式化百分比指标
# ============================================================

# 注意：
# 这里没有改变原始数值，只是后面Excel展示时再处理。
# CTR、CVR、ACOS、Ad_Order_Rate本质上都是小数。


# ============================================================
# 第十一部分：筛选“广告烧钱”SKU
# ============================================================

# 这里暂时使用：
# ACOS > 30%
#
# 0.30 = 30%

bad_ad = df[
    df["ACOS"] > 0.30
].copy()


print("\n" + "=" * 60)
print("第五阶段：筛选高ACOS SKU")
print("=" * 60)

print("高ACOS SKU数量：")
print(len(bad_ad))


# 按ACOS从高到低排序

bad_ad = bad_ad.sort_values(
    "ACOS",
    ascending=False
)


print("\n高ACOS SKU：")

print(
    bad_ad[
        [
            "SKU",
            "Product_Name",
            "Sales",
            "Ad_Spend",
            "ACOS",
            "ROAS",
            "Orders",
            "Ad_Orders",
            "Ad_Order_Rate"
        ]
    ]
)


# ============================================================
# 第十二部分：自动诊断SKU问题
# ============================================================

def diagnose_sku(row):

    # --------------------------------------------------------
    # 第一类：
    # 高ACOS + 高CPC
    # --------------------------------------------------------

    if row["ACOS"] > 0.30 and row["CPC"] > 10:

        return "高ACOS+高CPC：重点检查关键词竞价"


    # --------------------------------------------------------
    # 第二类：
    # 高ACOS + 低CVR
    # --------------------------------------------------------

    elif row["ACOS"] > 0.30 and row["CVR"] < 0.08:

        return "高ACOS+低CVR：重点检查Listing转化"


    # --------------------------------------------------------
    # 第三类：
    # 高ACOS + 低CTR
    # --------------------------------------------------------

    elif row["ACOS"] > 0.30 and row["CTR"] < 0.01:

        return "高ACOS+低CTR：重点检查主图和广告相关性"


    # --------------------------------------------------------
    # 第四类：
    # 只有ACOS偏高
    # --------------------------------------------------------

    elif row["ACOS"] > 0.30:

        return "ACOS偏高：需要进一步检查广告结构"


    # --------------------------------------------------------
    # 第五类：
    # 正常
    # --------------------------------------------------------

    else:

        return "正常"


# 对每一个SKU进行诊断

df["Diagnosis"] = df.apply(
    diagnose_sku,
    axis=1
)


# ============================================================
# 第十三部分：自动生成运营建议
# ============================================================

def operation_action(diagnosis):

    # 高CPC
    if diagnosis == "高ACOS+高CPC：重点检查关键词竞价":

        return (
            "检查高花费低转化关键词，"
            "降低无效词Bid，优化关键词投放"
        )


    # 低CVR
    elif diagnosis == "高ACOS+低CVR：重点检查Listing转化":

        return (
            "检查主图、价格、Review、优惠券、"
            "五点描述和Listing转化"
        )


    # 低CTR
    elif diagnosis == "高ACOS+低CTR：重点检查主图和广告相关性":

        return (
            "检查主图、标题、关键词相关性及广告投放词"
        )


    # 普通高ACOS
    elif diagnosis == "ACOS偏高：需要进一步检查广告结构":

        return (
            "进一步检查广告活动、关键词、"
            "搜索词及预算分配"
        )


    # 正常
    else:

        return "暂不处理，持续监控"


# 根据诊断结果生成运营建议

df["Action"] = df["Diagnosis"].apply(
    operation_action
)


# ============================================================
# 第十四部分：查看自动诊断结果
# ============================================================

print("\n" + "=" * 60)
print("第六阶段：SKU自动诊断结果")
print("=" * 60)


diagnosis_result = df[
    [
        "SKU",
        "Product_Name",
        "CTR",
        "CVR",
        "CPC",
        "ACOS",
        "ROAS",
        "Ad_Order_Rate",
        "Diagnosis",
        "Action"
    ]
]


print(diagnosis_result.head(20))


# ============================================================
# 第十五部分：只查看异常SKU
# ============================================================

bad_sku = df[
    df["Diagnosis"] != "正常"
].copy()


bad_sku = bad_sku.sort_values(
    "ACOS",
    ascending=False
)


print("\n" + "=" * 60)
print("第七阶段：异常SKU诊断报告")
print("=" * 60)


print(
    bad_sku[
        [
            "SKU",
            "Product_Name",
            "CTR",
            "CVR",
            "CPC",
            "ACOS",
            "ROAS",
            "Ad_Order_Rate",
            "Diagnosis",
            "Action"
        ]
    ]
)


# ============================================================
# 第十六部分：生成Excel报告
# ============================================================

# 创建ExcelWriter
with pd.ExcelWriter(
    excel_path,
    engine="openpyxl"
) as writer:

    # --------------------------------------------------------
    # Sheet 1：全部SKU
    # --------------------------------------------------------

    df.to_excel(
        writer,
        sheet_name="全部SKU",
        index=False
    )


    # --------------------------------------------------------
    # Sheet 2：高ACOS SKU
    # --------------------------------------------------------

    bad_ad.to_excel(
        writer,
        sheet_name="高ACOS SKU",
        index=False
    )


    # --------------------------------------------------------
    # Sheet 3：异常诊断
    # --------------------------------------------------------

    bad_sku[
        [
            "SKU",
            "Product_Name",
            "CTR",
            "CVR",
            "CPC",
            "ACOS",
            "ROAS",
            "Ad_Order_Rate",
            "Diagnosis",
            "Action"
        ]
    ].to_excel(
        writer,
        sheet_name="SKU诊断报告",
        index=False
    )


# ============================================================
# 第十七部分：程序结束
# ============================================================

print("\n" + "=" * 60)
print("项目运行完成！")
print("=" * 60)

print("\n原始CSV：")
print(csv_path)

print("\n最终Excel：")
print(excel_path)

print("\n最终SKU数量：")
print(len(df))

print("\n异常SKU数量：")
print(len(bad_sku))

print("\nExcel报告已经成功生成。")