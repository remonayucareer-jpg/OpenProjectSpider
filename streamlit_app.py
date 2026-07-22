from datetime import datetime

import streamlit as st
import pandas as pd

from openproject_exporter import (
    DEFAULT_START_DATE,
    EXPORT_COLUMNS,
    build_excel_bytes,
    build_dataframe,
    make_output_filename,
)


st.set_page_config(page_title="OpenProject AI运营工单导出", layout="wide")

st.title("OpenProject AI运营工单导出")
st.caption("选择工单创建日期范围，系统会抓取 AI运营工单并生成 Excel。")

default_date = datetime.strptime(DEFAULT_START_DATE, "%Y-%m-%d").date()


def get_secret_value(name):
    try:
        return st.secrets.get(name, None)
    except Exception:
        return None


secret_api_key = get_secret_value("OPENPROJECT_API_KEY")
secret_auth_header = get_secret_value("OPENPROJECT_AUTH_HEADER")

with st.sidebar:
    st.header("连接设置")
    typed_api_key = st.text_input(
        "OpenProject API Key",
        type="password",
        disabled=bool(secret_api_key or secret_auth_header),
        help="部署到 Streamlit Cloud 后建议放在 Secrets 里；本地临时测试可以在这里填写。",
    )
    if secret_api_key or secret_auth_header:
        st.success("已从 Secrets 读取鉴权信息。")
    else:
        st.info("本地测试时可临时填写 API Key；上传 GitHub 时不要把 Key 写进代码。")

    st.divider()
    st.header("过滤设置")
    exclude_keywords_input = st.text_input(
        "排除主题关键词（逗号分隔）",
        value="",
        placeholder="例如：测试,硬件实施",
        help="输入关键词，工单主题（subject）中包含这些关键词的行将被过滤掉，不展示也不导出。多个关键词用逗号分隔。",
    )
    # 解析排除关键词（支持中文逗号、英文逗号）
    exclude_keywords = []
    if exclude_keywords_input:
        # 先将中文逗号替换为英文逗号，再按英文逗号分割
        normalized = exclude_keywords_input.replace("，", ",")
        exclude_keywords = [kw.strip() for kw in normalized.split(",") if kw.strip()]
    if exclude_keywords:
        st.info(f"已设置 {len(exclude_keywords)} 个排除关键词：{'、'.join(exclude_keywords)}")

with st.form("export_form"):
    col1, col2 = st.columns(2)
    with col1:
        start_date = st.date_input("开始日期", value=default_date)
    with col2:
        end_date = st.date_input("结束日期", value=default_date)

    submitted = st.form_submit_button("抓取并预览数据", use_container_width=True)

if submitted:
    if end_date < start_date:
        st.error("结束日期不能早于开始日期。")
    else:
        api_key = secret_api_key or typed_api_key.strip() or None
        auth_header = secret_auth_header

        if not api_key and not auth_header:
            st.error("请先在左侧填写 OpenProject API Key，或在 Streamlit Secrets 中配置。")
            st.stop()

        try:
            with st.spinner("正在抓取工单，请稍等..."):
                # 只抓取一次，同时返回过滤后数据、原始数据和原始总数
                df_filtered, df_all, total_count = build_dataframe(
                    start_date,
                    end_date,
                    api_key=api_key,
                    auth_header=auth_header,
                    exclude_keywords=exclude_keywords if exclude_keywords else None,
                )
                filtered_count = len(df_filtered)
                excluded_count = total_count - filtered_count

            # 显示统计信息
            col_stat1, col_stat2, col_stat3 = st.columns(3)
            with col_stat1:
                st.metric("原始数据总量", total_count)
            with col_stat2:
                st.metric("过滤后数据量", filtered_count)
            with col_stat3:
                st.metric("因主题过滤排除", excluded_count)

            # 如果有排除，显示被排除的工单详情
            if excluded_count > 0 and exclude_keywords:
                excluded_df = df_all[~df_all.index.isin(df_filtered.index)]
                with st.expander(f"查看被排除的 {excluded_count} 条工单详情"):
                    st.dataframe(
                        excluded_df[["工单编号", "酒店名称", "工单状态", "工单创建时间"]],
                        use_container_width=True,
                        hide_index=True,
                    )

            # 预览过滤后的数据表格
            st.subheader("数据预览")
            st.dataframe(
                df_filtered,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "具体需求": st.column_config.TextColumn("具体需求", width="large"),
                },
            )

            # 下载按钮
            st.divider()
            col_dl1, col_dl2 = st.columns([1, 3])
            with col_dl1:
                # 生成过滤后的 Excel
                excel_bytes, row_count = build_excel_bytes(
                    start_date,
                    end_date,
                    api_key=api_key,
                    auth_header=auth_header,
                    exclude_keywords=exclude_keywords if exclude_keywords else None,
                )
                filename = make_output_filename(start_date, end_date)
                st.download_button(
                    "📥 下载过滤后的 Excel",
                    data=excel_bytes,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            with col_dl2:
                st.caption(f"下载的 Excel 包含 {row_count} 条数据（已排除主题含有关键词的工单）")

        except Exception as exc:
            st.error(f"生成失败：{exc}")
            st.exception(exc)

with st.expander("导出字段说明"):
    st.write("、".join(EXPORT_COLUMNS))
