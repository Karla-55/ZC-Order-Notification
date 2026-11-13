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

# 仓库关键词（纯大写，匹配时会先清洗原数据）
WH_INCLUDE = ["NORA", "OSA"]
WH_EXCLUDE = ["ROSA"]
# ==============================

# 辅助函数：清洗文本，只保留字母数字并转大写
def clean_text(s):
    if pd.isna(s):
        return ""
    s = str(s).strip().upper()
    s = re.sub(r'[^A-Z0-9]', '', s)  # 移除非字母数字的所有字符
    return s

uploaded_file = st.file_uploader("上传原始订单Excel", type=["xlsx", "xls"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, dtype=str)
        st.success(f"上传成功，原始数据共 {len(df)} 条")

        # 步骤2：删除指定列
        exist_cols = [c for c in COLS_TO_DELETE if c in df.columns]
        if exist_cols:
            df = df.drop(columns=exist_cols)

        # 步骤4(1)：删除无效Email
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

        df_exp = pd.DataFrame()
        df_ltl = pd.DataFrame()
        df_osa = pd.DataFrame()

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            # Sheet1：总订单
            df.to_excel(writer, sheet_name="总订单", index=False)

            # Sheet2：ZC-Fedex&UPS
            if "快递类型" in df.columns:
                express_col = df["快递类型"].astype(str).str.strip().str.lower()
                exp_mask = express_col.str.contains("|".join(EXPRESS_KEYWORDS), case=False, na=False)
                df_exp = df[exp_mask].reset_index(drop=True)
                df_exp.to_excel(writer, sheet_name="ZC-Fedex&UPS", index=False)

                # Sheet3：ZC-LTL
                pickup_mask = express_col.str.contains(SELF_PICKUP_KEYWORD, case=False, na=False)
                ltl_mask = ~(exp_mask | pickup_mask)
                df_ltl = df[ltl_mask].reset_index(drop=True)
                df_ltl.to_excel(writer, sheet_name="ZC-LTL", index=False)
            else:
                st.warning("未找到「快递类型」列，跳过快递分类筛选")

            # Sheet4：OSA&NORA（增强清洗后匹配）
            if "仓库编号" in df.columns:
                # 生成清洗后的临时列用于匹配
                wh_clean = df["仓库编号"].apply(clean_text)
                
                # 包含 NORA 或 OSA
                keep_mask = pd.Series(False, index=df.index)
                for kw in WH_INCLUDE:
                    keep_mask = keep_mask | wh_clean.str.contains(kw, na=False)
                
                # 排除包含 ROSA 的
                remove_mask = pd.Series(False, index=df.index)
                for kw in WH_EXCLUDE:
                    remove_mask = remove_mask | wh_clean.str.contains(kw, na=False)
                
                final_mask = keep_mask & ~remove_mask
                df_osa = df[final_mask].reset_index(drop=True)
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
        col1.metric("总订单", len(df))
        col2.metric("Fedex&UPS", len(df_exp))
        col3.metric("LTL", len(df_ltl))
        col4.metric("OSA&NORA", len(df_osa))

        # ========== 调试面板：仓库编号匹配详情 ==========
        if "仓库编号" in df.columns:
            with st.expander("🔍 仓库编号匹配调试（点击展开）", expanded=False):
                st.caption("以下是总订单中所有出现过的仓库编号，以及是否被命中 OSA&NORA")
                # 统计每个唯一值的出现次数和命中状态
                debug_df = pd.DataFrame({
                    "原始仓库编号": df["仓库编号"].fillna("(空值)").unique()
                })
                debug_df["清洗后"] = debug_df["原始仓库编号"].apply(clean_text)
                debug_df["是否命中OSA&NORA"] = debug_df["清洗后"].apply(
                    lambda x: any(kw in x for kw in WH_INCLUDE) 
                              and not any(kw in x for kw in WH_EXCLUDE)
                )
                # 统计出现次数
                count_map = df["仓库编号"].fillna("(空值)").value_counts().to_dict()
                debug_df["出现次数"] = debug_df["原始仓库编号"].map(count_map)
                debug_df = debug_df.sort_values("出现次数", ascending=False).reset_index(drop=True)
                
                st.dataframe(debug_df, use_container_width=True, height=400)
                st.caption("如果有你认为应该命中但显示「否」的，请把原始仓库编号发给我，我补充匹配规则")

    except Exception as e:
        st.error(f"处理出错：{e}")
        st.info("请检查表头列名是否与规则完全一致")
