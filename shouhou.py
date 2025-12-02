# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

st.set_page_config(page_title="多平台退款/退货数据清洗工具 v3", layout="wide")
st.title("📦 多平台退款/退货数据清洗工具 v3（最终稳定版）")

uploaded_files = st.file_uploader(
    "请上传所有相关文件（可多选）",
    type=["xlsx", "xls", "csv"],
    accept_multiple_files=True
)

# ========= 通用函数（处理订单号） =========
def normalize_order_id(series):
    """
    将订单号强制转为纯字符串：
    - 去除逗号千分位
    - 去除 .0
    - 去除空格
    """
    return (
        series.astype(str)
            .str.replace(",", "", regex=False)
            .str.replace(".0", "", regex=False)
            .str.replace(" ", "", regex=False)
            .str.strip()
    )

# ========= 第一类：Amazon 买家退货 reason 映射 =========
amazon_reason_mapping = {
    "UNWANTED_ITEM": "不想要的商品",
    "DEFECTIVE": "商品存在瑕疵",
    "NOT_AS_DESCRIBED": "和网站上的描述不一致",
    "SWITCHEROO": "亚马逊发了错误的产品",
    "MISSED_ESTIMATED_DELIVERY": "超过预期时间未交付",
    "MISSING_PARTS": "配送中商品或配件丢失",
    "FOUND_BETTER_PRICE": "发现更优惠的价格",
    "DAMAGED_BY_FC": "商品运送到时存在残损或瑕疵",
    "QUALITY_UNACCEPTABLE": "商品性能或质量未达预期",
    "ORDERED_WRONG_ITEM": "买错货",
    "UNDELIVERABLE_REFUSED": "无法配送_已拒收",
    "DAMAGED_BY_CARRIER": "商品运送到时存在残损或瑕疵",
    "UNAUTHORIZED_PURCHASE": "未经授权购买：例如欺诈",
    "NEVER_ARRIVED": "未配送",
    "UNDELIVERABLE_UNKNOWN": "无法配送_未知原因",
    "NO_REASON_GIVEN": "没有理由",
    "EXTRA_ITEM": "货件中包含其他商品",
    "NOT_COMPATIBLE": "商品与当前系统不兼容",
    "APPAREL_STYLE": "不喜欢产品外观风格/款式",
    "UNDELIVERABLE_INSUFFICIENT_ADDRESS": "无法配送_地址无效",
    "APPAREL_TOO_SMALL": "产品外观太小",
    "APPAREL_TOO_LARGE": "产品外观太大",
    "MISORDERED": "订购错误的款式/尺寸/颜色",
    "UNDELIVERABLE_CARRIER_MISS_SORTED": "无法交付_承运人丢失",
    "UNDELIVERABLE_FAILED_DELIVERY_ATTEMPTS": "无法配送_尝试配送失败",
    "UNDELIVERABLE_MISSING_LABEL": "无法交付_丢失标签",
    "UNDELIVERABLE_UNCLAIMED": "无法配送_无人认领",
    "PERFORMANCE/QUALITY NOT UP TO EXPECTATIONS": "商品性能或质量未达预期",
    "DAMAGED/DEFECTIVE ON ARRIVAL": "商品运送到时存在残损或瑕疵",
    "MISSING ITEMS OR ACCESSORIES": "配送中商品或配件丢失",
    "UNWANTED ITEM": "不想要的商品",
    "WRONG_SIZE": "尺寸错误",
    "MISSED ESTIMATED DELIVERY": "超过预期时间未交付",
    "ORDERED WRONG ITEM": "买错货",
    "NO REASON GIVEN": "没有理由",
    "UNDELIVERABLE UNKNOWN": "无法配送_未知原因",
    "UNDELIVERABLE REFUSED": "无法配送_已拒收",
    "UNAUTHORIZED PURCHASE": "未经授权购买：例如欺诈",
    "UNDELIVERABLE FAILED DELIVERY ATTEMPTS": "无法配送_尝试配送失败",
    "WRONG ITEM SHIPPED": "亚马逊发了错误的产品",
    "FOUND BETTER PRICE ELSEWHERE": "发现更优惠的价格",
    "NOT AS DESCRIBED ON WEBSITE": "和网站上的描述不一致",
    "DAMAGED/DEFECTIVE AFTER ARRIVAL": "商品运送到时存在残损或瑕疵",
    "NOT COMPATIBLE WITH EXISTING SYSTEM": "商品与当前系统不兼容",
    "EXTRA ITEM INCLUDED IN SHIPMENT": "货件中包含其他商品"
}

# ========= 第三类：Amazon换货 reason mapping =========
amazon_exchange_reason = {
    "0": "其他",
    "1": "丢失",
    "2": "存在缺陷",
    "3": "配送过程中残损",
    "4": "商品配送错误",
    "5": "商品在配送过程中丢失",
    "6": "发货人丢失商品",
    "7": "目录错误/买错商品",
    "8": "配送到错误的地址",
    "9": "配送问题（地址正确）",
    "10": "DC/FC处理中心残损",
    "11": "未收到商品",
    "12": "政策例外/买家错误"
}

# ========== 辅助函数 ==========
def safe_get(df, col):
    for c in df.columns:
        if c.lower().strip() == col.lower().strip():
            return df[c]
    return pd.Series([None] * len(df))

def rename_safe(df, old, new):
    for c in df.columns:
        if c.lower().strip() == old.lower().strip():
            df.rename(columns={c: new}, inplace=True)

def extract_sku_from_temu(v):
    if pd.isna(v):
        return None
    parts = str(v).split("_")
    if len(parts) >= 3:
        return parts[1]
    return None

# ========= 主处理流程 =========
all_results = []

if uploaded_files:
    for file in uploaded_files:
        filename = file.name.lower()

        # ========== 第一类：Amazon买家退货 ==========
        if "amazon买家退货" in filename:
            df = pd.read_excel(file)
            df["order_id"] = normalize_order_id(safe_get(df, "order-id"))
            df["平台sku"] = safe_get(df, "平台sku")
            df["reason"] = safe_get(df, "reason").astype(str).str.upper().map(amazon_reason_mapping)
            df["platform"] = "Amazon"
            df["platform_refund_reason"] = df["platform"] + df["reason"]
            df["source_file"] = filename
            all_results.append(df[["order_id", "平台sku", "reason", "platform", "platform_refund_reason", "source_file"]])
            continue

        # ========== 第二类：Amazont退货报告 ==========
        if "amazont退货报告" in filename:
            df = pd.read_excel(file)
            rename_safe(df, "merchant_sku", "平台sku")
            rename_safe(df, "return_reason", "reason")
            rename_safe(df, "order_id", "order_id")
            df["order_id"] = normalize_order_id(df["order_id"])
            df["platform"] = "Amazon"
            df["platform_refund_reason"] = df["platform"] + df["reason"].astype(str)
            df["source_file"] = filename
            all_results.append(df[["order_id", "平台sku", "reason", "platform", "platform_refund_reason", "source_file"]])
            continue

        # ========== 第三类：Amazon后台换货 ==========
        if "平台amazon后台换货表" in filename:
            df = pd.read_excel(file)
            rename_safe(df, "sku", "平台sku")
            df["reason"] = safe_get(df, "replacement-reason-code").astype(str).map(amazon_exchange_reason)
            df["order_id"] = normalize_order_id(safe_get(df, "original-amazon-order-id"))
            df["platform"] = "Amazon"
            df["platform_refund_reason"] = df["platform"] + df["reason"]
            df["source_file"] = filename
            all_results.append(df[["order_id", "平台sku", "reason", "platform", "platform_refund_reason", "source_file"]])
            continue

        # ========== 第四类：Overstock后台退货单 ==========
        if "overstock后台退货单" in filename:
            df = pd.read_excel(file)
            rename_safe(df, "Return Reason Description", "reason")
            rename_safe(df, "Partner SKU", "平台sku")
            rename_safe(df, "Order Number", "order_id")
            df["order_id"] = normalize_order_id(df["order_id"])
            df["platform"] = "Overstock"
            df["platform_refund_reason"] = df["platform"] + df["reason"].astype(str)
            df["source_file"] = filename
            all_results.append(df[["order_id", "平台sku", "reason", "platform", "platform_refund_reason", "source_file"]])
            continue

        # ========== 第五类：TEMU后台退款表 ==========
        if "temu后台退款表" in filename:
            df = pd.read_excel(file)
            rename_safe(df, "售后原因", "reason")
            rename_safe(df, "订单编号", "order_id")
            df["order_id"] = normalize_order_id(df["order_id"])

            sku_col = [c for c in df.columns if "sku" in c.lower()][0]
            df["平台sku"] = df[sku_col].apply(extract_sku_from_temu)

            df["platform"] = "temu"
            df["platform_refund_reason"] = df["platform"] + df["reason"].astype(str)
            df["source_file"] = filename

            all_results.append(df[["order_id", "平台sku", "reason", "platform", "platform_refund_reason", "source_file"]])
            continue

        # ========== 第六类：Tiktok后台退款表 ==========
        if "tiktok后台退款表" in filename:
            df = pd.read_excel(file)
            rename_safe(df, "Order ID", "order_id")
            rename_safe(df, "Seller SKU", "平台sku")
            rename_safe(df, "Return Reason", "reason")

            df["order_id"] = normalize_order_id(df["order_id"])
            df["platform"] = "Tiktok"
            df["platform_refund_reason"] = df["platform"] + df["reason"].astype(str)
            df["source_file"] = filename

            all_results.append(df[["order_id", "平台sku", "reason", "platform", "platform_refund_reason", "source_file"]])
            continue

        # ========== 第七类：VC退款核查 ==========
        if "vc退款核查" in filename:
            xls = pd.ExcelFile(file)

            # Orders sheet
            if "Orders下退款" in xls.sheet_names:
                df = pd.read_excel(xls, "Orders下退款")
                rename_safe(df, "SKU", "平台sku")
                rename_safe(df, "Order ID", "order_id")
                rename_safe(df, "Return Reason", "reason")

                df["order_id"] = normalize_order_id(df["order_id"])
                df["platform"] = "VC"
                df["platform_refund_reason"] = df["platform"] + df["reason"].astype(str)
                df["source_file"] = filename + "_orders"

                all_results.append(df[["order_id", "平台sku", "reason", "platform", "platform_refund_reason", "source_file"]])

            # Payments sheet
            if "Payments下退款" in xls.sheet_names:
                df2 = pd.read_excel(xls, "Payments下退款")
                rename_safe(df2, "Reason", "reason")
                rename_safe(df2, "Distributor Shipment Id", "order_id")

                df2["order_id"] = normalize_order_id(df2["order_id"])
                df2["平台sku"] = None
                df2["platform"] = "VC"
                df2["platform_refund_reason"] = df2["platform"] + df2["reason"].astype(str)
                df2["source_file"] = filename + "_payments"

                all_results.append(df2[["order_id", "平台sku", "reason", "platform", "platform_refund_reason", "source_file"]])

            continue

        # ========== 第八类：Walmart后台退款表 ==========
        if "walmart后台退款表" in filename:
            df = pd.read_excel(file)
            rename_safe(df, "RETURN_REASON", "reason")
            rename_safe(df, "CUSTOMER_ORDER_NO", "order_id")

            df["order_id"] = normalize_order_id(df["order_id"])
            df["平台sku"] = None
            df["platform"] = "Walmart"
            df["platform_refund_reason"] = df["platform"] + df["reason"].astype(str)
            df["source_file"] = filename

            all_results.append(df[["order_id", "平台sku", "reason", "platform", "platform_refund_reason", "source_file"]])
            continue

        # ========== 第九类：Wayfair后台退款表 ==========
        if "wayfair后台退款表" in filename:
            df = pd.read_excel(file)
            rename_safe(df, "原因", "reason")
            rename_safe(df, "SKU", "平台sku")

            po_col = [c for c in df.columns if "po" in c.lower()][0]
            df["order_id"] = normalize_order_id(df[po_col])

            df["platform"] = "Walmart"
            df["platform_refund_reason"] = df["platform"] + df["reason"].astype(str)
            df["source_file"] = filename

            all_results.append(df[["order_id", "平台sku", "reason", "platform", "platform_refund_reason", "source_file"]])
            continue


# ========== 合并输出 ==========
if all_results:
    final_df = pd.concat(all_results, ignore_index=True)

    st.subheader("🎉 清洗完成！预览前 20 行：")
    st.dataframe(final_df.head(20))

    buffer = BytesIO()
    final_df.to_excel(buffer, index=False)
    buffer.seek(0)

    st.download_button(
        "⬇️ 下载合并后的大表（Excel）",
        data=buffer,
        file_name="refund_merged_cleaned_v3.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
