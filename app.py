import pandas as pd
from datetime import datetime
import streamlit as st
from io import BytesIO

st.set_page_config(page_title="订单数据整理工具", layout="centered")
st.title("订单数据自动整理工具")

# ========== 规则配置 ==========
COLS_TO_DELETE = ["数据唯一编号列", "备注2"]

VALID_OP_GROUPS = [
    "HD-wellfor", "Lowes-wellfor", "HD-forclover", "Lowes-forclover",
    "Houzz", "Amazon&Walmart&Wayfair&Overstock", "HD-Angeleshome",
    "Target", "Lowes-Angeleshome", "Best buy",
    "Shopify-wellfor", "Shopify-forclover"
]

# 快递关键词（包含即匹配，不区分大小写）
EXPRESS_KEYWORDS = ["fedex", "ups"]
SELF_PICKUP_KEYWORD = "客户自提"
# ==============================

uploaded_file = st.file_uploader("上传原始订单Excel", type=["xlsx", "xls"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, dtype=str)
        st.success(f"上传成功，原始数据共 {len(df)} 条")

        # 步骤2：删除指定列
        exist_cols = [c for c in COLS_TO_DELETE if c in df.columns]
        if exist_cols:
            df = df.drop(columns=exist_cols)

        # 步骤4(1)：删除无效Email（空值、空白、N/A）
        if "Email" in df.columns:
            df["Email"] = df["Email"].astype(str).str.strip()
            invalid_email = (df["Email"].isna()) | (df["Email"] == "") | (df["Email"].str.upper() == "N/A")
            df = df[~invalid_email].reset_index(drop=True)
            st.info(f"已清理无效Email订单：{invalid_email.sum()} 条")
        else:
            st.warning("未找到「Email」列，跳过该步筛选")

        # 步骤4(2)：仅保留指定运营组
        if "运营组" in df.columns:
            df["运营组"] = df["运营组"].astype(str).str.strip()
            valid_mask = df["运营组"].isin(VALID_OP_GROUPS)
            df = df[valid_mask].reset_index(drop=True)
            st.info(f"已清理非指定运营组订单：{(~valid_mask).sum()} 条")
        else:
            st.warning("未找到「运营组」列，跳过该步筛选")

        # 生成Excel文件到内存
        output = BytesIO()
        today_str = datetime.now().strftime("%Y%m%d")
        filename = f"ZC-{today_str}.xlsx"

        # 初始化各Sheet数据（兜底空表，避免报错）
        df_exp = pd.DataFrame()
        df_ltl = pd.DataFrame()
        df_osa = pd.DataFrame()

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            # Sheet1：总订单
            df.to_excel(writer, sheet_name="总订单", index=False)

            # Sheet2：ZC-Fedex&UPS（关键词包含匹配，不区分大小写）
            if "快递类型" in df.columns:
                express_col = df["快递类型"].astype(str).str.strip().str.lower()
                exp_mask = express_col.str.contains("|".join(EXPRESS_KEYWORDS), case=False, na=False)
                df_exp = df[exp_mask].reset_index(drop=True)
                df_exp.to_excel(writer, sheet_name="ZC-Fedex&UPS", index=False)

                # Sheet3：ZC-LTL（排除Fedex、UPS、客户自提）
                pickup_mask = express_col.str.contains(SELF_PICKUP_KEYWORD, case=False, na=False)
                ltl_mask = ~(exp_mask | pickup_mask)
                df_ltl = df[ltl_mask].reset_index(drop=True)
                df_ltl.to_excel(writer, sheet_name="ZC-LTL", index=False)
            else:
                st.warning("未找到「快递类型」列，跳过快递分类筛选")

            # Sheet4：OSA&NORA（包含NORA/OSA，排除ROSA，不区分大小写）
            if "仓库编号" in df.columns:
                wh_col = df["仓库编号"].astype(str).str.strip().str.upper()
                # 包含 NORA 或 OSA
                keep_mask = wh_col.str.contains("NORA", na=False) | wh_col.str.contains("OSA", na=False)
                # 排除包含 ROSA 的
                remove_mask = wh_col.str.contains("ROSA", na=False)
                final_mask = keep_mask & ~remove_mask
                df_osa = df[final_mask].reset_index(drop=True)
                df_osa.to_excel(writer, sheet_name="OSA&NORA", index=False)
            else:
                st.warning("未找到「仓库编号」列，跳过仓库筛选")

        output.seek(0)
        st.divider()
        st.success("✅ 数据整理完成")
        st.download_button(
            "下载整理后的Excel",
            data=output,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # 结果概览
        st.subheader("处理结果概览")
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("总订单", len(df))
        col2.metric("Fedex&UPS", len(df_exp))
        col3.metric("LTL", len(df_ltl))
        col4.metric("OSA&NORA", len(df_osa))

    except Exception as e:
        st.error(f"处理出错：{e}")
        st.info("请检查表头列名是否与规则完全一致")
