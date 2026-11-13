import pandas as pd
from datetime import datetime
import streamlit as st
from io import BytesIO
import re

st.set_page_config(page_title="订单数据整理工具", layout="wide")
st.title("订单数据自动整理工具")

# ========== 规则配置 ==========
COLS_TO_DELETE = ["数据唯一编号列", "备注2"]

VALID_OP_GROUPS = [
    "HD-wellfor", "Lowes-wellfor", "HD-forclover", "Lowes-forclover",
    "Houzz", "Amazon&Walmart&Wayfair&Overstock", "HD-Angeleshome",
    "Target", "Lowes-Angeleshome", "Best buy",
    "Shopify-wellfor", "Shopify-forclover"
]

# 快递关键词
EXPRESS_KEYWORDS = ["fedex", "ups"]
SELF_PICKUP_KEYWORD = "客户自提"

# 仓库关键词
WH_INCLUDE = ["NORA", "OSA"]
WH_EXCLUDE = ["ROSA"]
# ==============================

# 辅助函数：清洗文本，只保留字母数字并转大写
def clean_text(s):
    if pd.isna(s):
        return ""
    s = str(s).strip().upper()
    s = re.sub(r'[^A-Z0-9]', '', s)
    return s

uploaded_file = st.file_uploader("上传原始订单Excel", type=["xlsx", "xls"])

if uploaded_file:
    try:
        df_raw = pd.read_excel(uploaded_file, dtype=str)
        st.success(f"上传成功，原始数据共 {len(df_raw)} 条")

        # 步骤2：删除指定列（所有Sheet共用的基础清洗）
        exist_cols = [c for c in COLS_TO_DELETE if c in df_raw.columns]
        if exist_cols:
            df_base = df_raw.drop(columns=exist_cols)
        else:
            df_base = df_raw.copy()

        # ========== 总订单、快递、LTL 专用：先过滤运营组，再过滤Email ==========
        df_op_filtered = df_base.copy()
        if "运营组" in df_op_filtered.columns:
            df_op_filtered["运营组"] = df_op_filtered["运营组"].astype(str).str.strip()
            valid_op_mask = df_op_filtered["运营组"].isin(VALID_OP_GROUPS)
            df_op_filtered = df_op_filtered[valid_op_mask].reset_index(drop=True)
            st.info(f"总订单/快递/LTL：已清理非指定运营组订单 {(~valid_op_mask).sum()} 条")
        else:
            st.warning("未找到「运营组」列，跳过运营组筛选")

        df_total = df_op_filtered.copy()
        if "Email" in df_total.columns:
            df_total["Email"] = df_total["Email"].astype(str).str.strip()
            invalid_email = (df_total["Email"].isna()) | (df_total["Email"] == "") | (df_total["Email"].str.upper() == "N/A")
            df_total = df_total[~invalid_email].reset_index(drop=True)
            st.info(f"总订单/快递/LTL：已清理无效Email订单 {invalid_email.sum()} 条")
        else:
            st.warning("未找到「Email」列，跳过Email筛选")

        # ========== OSA&NORA 专用：不做运营组过滤，不做Email过滤，仅按仓库筛选 ==========
        st.info("OSA&NORA：不筛选运营组，不筛选Email，仅按仓库编号匹配")

        # 生成Excel文件到内存
        output = BytesIO()
        today_str = datetime.now().strftime("%Y%m%d")
        filename = f"ZC-{today_str}.xlsx"

        df_exp = pd.DataFrame()
        df_ltl = pd.DataFrame()
        df_osa = pd.DataFrame()

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            # Sheet1：总订单（已过滤运营组+Email）
            df_total.to_excel(writer, sheet_name="总订单", index=False)

            # Sheet2：ZC-Fedex&UPS（基于总订单）
            if "快递类型" in df_total.columns:
                express_col = df_total["快递类型"].astype(str).str.strip().str.lower()
                exp_mask = express_col.str.contains("|".join(EXPRESS_KEYWORDS), case=False, na=False)
                df_exp = df_total[exp_mask].reset_index(drop=True)
                df_exp.to_excel(writer, sheet_name="ZC-Fedex&UPS", index=False)

                # Sheet3：ZC-LTL（基于总订单）
                pickup_mask = express_col.str.contains(SELF_PICKUP_KEYWORD, case=False, na=False)
                ltl_mask = ~(exp_mask | pickup_mask)
                df_ltl = df_total[ltl_mask].reset_index(drop=True)
                df_ltl.to_excel(writer, sheet_name="ZC-LTL", index=False)
            else:
                st.warning("未找到「快递类型」列，跳过快递分类筛选")

            # Sheet4：OSA&NORA（基于最原始的删列数据，不过滤运营组、不过滤Email）
            if "仓库编号" in df_base.columns:
                wh_clean = df_base["仓库编号"].apply(clean_text)
                
                # 包含 NORA 或 OSA
                keep_mask = pd.Series(False, index=df_base.index)
                for kw in WH_INCLUDE:
                    keep_mask = keep_mask | wh_clean.str.contains(kw, na=False)
                
                # 排除包含 ROSA 的
                remove_mask = pd.Series(False, index=df_base.index)
                for kw in WH_EXCLUDE:
                    remove_mask = remove_mask | wh_clean.str.contains(kw, na=False)
                
                final_mask = keep_mask & ~remove_mask
                df_osa = df_base[final_mask].reset_index(drop=True)
                df_osa.to_excel(writer, sheet_name="OSA&NORA", index=False)
            else:
                st.warning("未找到「仓库编号」列，跳过仓库筛选")

        output.seek(0)
        st.divider()
        st.success("✅ 数据整理完成")
        st.download_button(
            "📥 下载整理后的Excel",
            data=output,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # 结果概览
        st.subheader("处理结果概览")
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("总订单", len(df_total))
        col2.metric("Fedex&UPS", len(df_exp))
        col3.metric("LTL", len(df_ltl))
        col4.metric("OSA&NORA", len(df_osa))

        # 调试面板：仓库编号匹配详情（基于最原始的删列数据）
        if "仓库编号" in df_base.columns:
            with st.expander("🔍 仓库编号匹配调试（点击展开）", expanded=False):
                st.caption("统计范围：已删除指定列，未过滤运营组、未过滤Email")
                debug_df = pd.DataFrame({
                    "原始仓库编号": df_base["仓库编号"].fillna("(空值)").unique()
                })
                debug_df["清洗后"] = debug_df["原始仓库编号"].apply(clean_text)
                debug_df["是否命中OSA&NORA"] = debug_df["清洗后"].apply(
                    lambda x: any(kw in x for kw in WH_INCLUDE) 
                              and not any(kw in x for kw in WH_EXCLUDE)
                )
                count_map = df_base["仓库编号"].fillna("(空值)").value_counts().to_dict()
                debug_df["出现次数"] = debug_df["原始仓库编号"].map(count_map)
                debug_df = debug_df.sort_values("出现次数", ascending=False).reset_index(drop=True)
                
                st.dataframe(debug_df, use_container_width=True, height=400)

    except Exception as e:
        st.error(f"处理出错：{e}")
        st.info("请检查表头列名是否与规则完全一致")
