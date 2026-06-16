from __future__ import annotations

from datetime import datetime

import streamlit as st

from processor import export_summary_workbook, infer_month_label, process_files


st.set_page_config(page_title="出口数据汇整", page_icon="📊", layout="wide")

st.markdown(
    """
    <style>
    .block-container { padding-top: 2rem; max-width: 1280px; }
    div[data-testid="stMetric"] {
        border: 1px solid #d8dee8;
        border-radius: 8px;
        padding: 14px 16px;
        background: #f8fafc;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("出口数据汇整")
st.caption("上传大连和铁岭 Excel，自动生成香港/新加坡外销汇整表。")

left, right = st.columns(2)
with left:
    dalian_file = st.file_uploader("大连文件", type=["xlsx"], key="dalian")
with right:
    tieling_file = st.file_uploader("铁岭文件", type=["xlsx"], key="tieling")

default_month = infer_month_label(
    getattr(dalian_file, "name", None),
    getattr(tieling_file, "name", None),
)
month_label = st.text_input("输出月份", value=default_month)

if not dalian_file or not tieling_file:
    st.info("请上传两个文件后生成汇整表。")
    st.stop()

try:
    tables = process_files(dalian_file, tieling_file)
except Exception as exc:
    st.error(f"处理失败：{exc}")
    st.stop()

metric_cols = st.columns(4)
metric_cols[0].metric("大连明细", len(tables.dalian))
metric_cols[1].metric("铁岭明细", len(tables.tieling))
metric_cols[2].metric("汇整行数", len(tables.final))
metric_cols[3].metric("总销量（吨）", f"{tables.final.loc[tables.final['产品名称'] != '小计', '月销量（吨）'].sum():.4g}")

tab_final, tab_dalian, tab_tieling = st.tabs(["汇整预览", "大连处理结果", "铁岭处理结果"])
with tab_final:
    st.dataframe(tables.final, use_container_width=True, hide_index=True)
with tab_dalian:
    st.dataframe(tables.dalian, use_container_width=True, hide_index=True)
with tab_tieling:
    st.dataframe(tables.tieling, use_container_width=True, hide_index=True)

workbook_bytes = export_summary_workbook(tables.final, month_label=month_label)
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
st.download_button(
    "下载汇整 Excel",
    data=workbook_bytes,
    file_name=f"{month_label}出口数据汇整_香港&新加坡_{timestamp}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
