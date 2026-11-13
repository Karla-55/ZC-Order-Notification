import pandas as pd
from datetime import datetime
import streamlit as st
from io import BytesIO

st.set_page_config(page_title="订单数据整理工具", layout="centered")
st.title("订单数据自动整理工具")

# 规则配置
COLS_TO_DELETE = ["数据唯一编号列", "备注2"]
VALID_OP_GROUPS = [
    "HD-wellfor", "Lowes-wellfor", "HD-forclover", "Lowes-forclover",
    "Houzz", "Amazon&Walmart&Wayfair&Overstock", "HD-Angeleshome",
    "Target", "Lowes-Angeleshome", "Best buy",
    "Shopify-wellfor", "Shopify-forclover"
]
EXPRESS_LIST = ["Fedex", "UPS"]
LTL_EXCLUDE = EXPRESS_LIST + ["客户自提"]

uploaded_file = st.file_uploader("上传原始订单Excel", type=["xlsx", "xls"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, dtype=str)
        st.success(f"上传成功，原始数据共 {len(df)} 条")

        # 删除指定列
        exist_cols = [c for c in COLS_TO_DELETE if c in df.columns]
        if exist_cols:
            df = df.drop(columns=exist_cols)

        # 清洗无效Email
        if "Email" in df.columns:
            df["Email"] = df["Email"].astype(str).str.strip()
            invalid = (df["Email"].isna()) | (df["Email"] == "") | (df["Email"].str.upper() == "N/A")
            df = df[~invalid].reset_index(drop=True)

        # 筛选运营组
        if "运营组" in df.columns:
            df["运营组"] = df["运营组"].astype(str).str.strip()
            df = df[df["运营组"].isin(VALID_OP_GROUPS)].reset_index(drop=True)

        # 生成Excel到内存
        output = BytesIO()
        today_str = datetime.now().strftime("%Y%m%d")
        filename = f"ZC-{today_str}.xlsx"

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="总订单", index=False)

            if "快递类型" in df.columns:
                df_exp = df[df["快递类型"].astype(str).str.strip().isin(EXPRESS_LIST)]
                df_exp.to_excel(writer, sheet_name="ZC-Fedex&UPS", index=False)

                ltl_mask = ~df["快递类型"].astype(str).str.strip().isin(LTL_EXCLUDE)
                df_ltl = df[ltl_mask]
                df_ltl.to_excel(writer, sheet_name="ZC-LTL", index=False)

            if "仓库编号" in df.columns:
                wh = df["仓库编号"].astype(str).str.strip()
                mask = (wh.str.contains("NORA", case=False) | wh.str.contains("OSA", case=False)) \
                       & ~wh.str.contains("ROSA", case=False)
                df_osa = df[mask]
                df_osa.to_excel(writer, sheet_name="OSA&NORA", index=False)

        output.seek(0)
        st.divider()
        st.success("数据整理完成")
        st.download_button(
            "下载整理后的Excel",
            data=output,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # 数据概览
        st.subheader("处理结果概览")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("总订单", len(df))
        if "快递类型" in df.columns:
            c2.metric("Fedex&UPS", len(df_exp))
            c3.metric("LTL", len(df_ltl))
        if "仓库编号" in df.columns:
            c4.metric("OSA&NORA", len(df_osa))

    except Exception as e:
        st.error(f"处理出错：{e}")
        st.info("请检查表头列名是否与规则完全一致")
