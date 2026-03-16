import pandas as pd
import numpy as np
import streamlit as st
import re
import io
from datetime import datetime


def col_num_to_letter(n):
    """将 0-based 列索引转换为 Excel 列字母"""
    letters = []
    while n >= 0:
        letters.append(chr(65 + (n % 26)))
        n = n // 26 - 1
    return ''.join(reversed(letters))


def like_order_string(s):
    """判断一个字符串是否像订单号"""
    if not isinstance(s, str):
        return False
    if len(s) < 3 or len(s) > 50:
        return False
    if re.search('[\u4e00-\u9fff]', s):
        return False
    if re.match(r'^[A-Za-z0-9\-_/\.]+$', s):
        return True
    return False


def detect_column(df, sheet_desc):
    """
    自动识别订单号列，结合列名关键词和列内数据特征。
    返回 (推荐的列名, 所有列名按得分排序)
    假设 df 的列名已经过标准化（无空格）。
    """
    keywords = [
        "订单号", "订单编号", "订单", "编号",
        "order number", "order no", "orderno", "order id",
        "order", "id"
    ]

    kw_scores = {}
    for col in df.columns:
        col_lower = str(col).lower()
        score = 0
        for kw in keywords:
            if kw.lower() in col_lower:
                score += 1
        kw_scores[col] = score

    content_scores = {}
    sample_size = min(100, len(df))
    sample = df.head(sample_size)
    for col in df.columns:
        non_null = sample[col].notna().sum()
        if non_null == 0:
            content_scores[col] = 0.0
            continue
        values = sample[col].dropna().astype(str).tolist()
        like_count = sum(1 for v in values if like_order_string(v))
        content_scores[col] = like_count / non_null

    total_scores = {}
    for col in df.columns:
        kw = kw_scores[col]
        content = content_scores[col]
        if kw > 0:
            total = kw * 10 + content
        else:
            total = content
        total_scores[col] = total

    sorted_columns = sorted(df.columns.tolist(), key=lambda c: total_scores[c], reverse=True)
    max_score = max(total_scores.values()) if total_scores else 0
    recommended = sorted_columns[0] if sorted_columns and max_score > 0 else None
    return recommended, sorted_columns


def clean_dataframe(df, exclude_columns=None):
    """
    清洗DataFrame：对非排除列进行向前填充，排除列保持原始值不变。
    """
    if exclude_columns is None:
        exclude_columns = []
    # 用 repr 打印排除列列表，以便查看精确字符串（含空格、括号等）
    st.code(f"清洗函数收到的排除列列表 (repr): {repr(exclude_columns)}")
    df_clean = df.copy()
    for col in df_clean.columns:
        if col not in exclude_columns:
            # 对非排除列：将空字符串转为 NaN，向前填充，再将 NaN 转回空字符串
            df_clean[col] = df_clean[col].replace(r'^\s*$', np.nan, regex=True)
            df_clean[col] = df_clean[col].ffill()
            df_clean[col] = df_clean[col].fillna('')
    return df_clean


def safe_order_str(x):
    """对订单号字符串进行简单清理：去除首尾空格。"""
    if pd.isna(x):
        return ''
    return str(x).strip()


def is_numeric_column(col_name):
    """
    判断某列是否应该转换为数值类型（金额、数量等）。
    """
    numeric_keywords = [
        "总价", "金额", "单价", "运费", "改价", "实付款", "结算价",
        "price", "amount", "freight", "payment", "settlement",
        "order amount", "shipping fee", "discount amount", "unit price",
        "total amount", "fee", "discount", "shipping",
        "initial payment", "balance payment", "tax",
        "数量", "quantity", "qty"
    ]
    col_lower = str(col_name).lower()
    for kw in numeric_keywords:
        if kw.lower() in col_lower:
            return True
    return False


def convert_numeric_columns(df):
    """
    将DataFrame中应转为数值的列转换为数值类型（float），
    无法转换的变为 NaN。
    """
    df_numeric = df.copy()
    for col in df_numeric.columns:
        if is_numeric_column(col):
            df_numeric[col] = pd.to_numeric(df_numeric[col], errors='coerce')
    return df_numeric


def main():
    st.set_page_config(
        page_title="订单匹配工具",
        page_icon="🔗",
        layout="wide",
        initial_sidebar_state="expanded"
    )

    with st.sidebar:
        st.header("📖 使用说明")
        st.markdown("""
        ### 操作步骤：
        1. **上传汇总表** 📊
        2. **上传明细表** 📋
        3. **选择订单号列** 🏷️
        4. **开始匹配** 🔗
        5. **下载结果** ⬇️

        ---
        ### 功能说明：
        ✅ 自动识别订单号列  
        ✅ 支持手动选择列  
        ✅ 明细表智能清洗：可手动指定不填充的列（如金额列），避免数值重复  
        ✅ 可下载清洗后明细表（订单号列自动转文本）  
        ✅ 清洗前后订单号对比，便于核对  
        ✅ 匹配结果中金额列、数量列自动转换为数值，无需手动修改格式  
        ✅ 汇总表/明细表数据预览，直观选择列  
        """)
        st.markdown("---")
        st.markdown("### 版本信息")
        st.info("版本: v1.18.0（增强排除列调试）")

    st.title("🔗 订单匹配工具")
    st.markdown("根据订单号匹配汇总表和明细表数据")
    st.markdown("---")

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("📊 汇总表")
        summary_file = st.file_uploader("上传汇总信息 Excel 文件", type=['xlsx', 'xls'], key="summary_file")
        if summary_file:
            st.success(f"✅ 已上传: {summary_file.name}")

    with col2:
        st.subheader("📋 明细表")
        detail_file = st.file_uploader("上传明细信息 Excel 文件", type=['xlsx', 'xls'], key="detail_file")
        if detail_file:
            st.success(f"✅ 已上传: {detail_file.name}")

    df_summary = None
    df_detail = None
    df_detail_raw = None
    summary_col = None
    detail_col = None

    if summary_file:
        try:
            df_summary = pd.read_excel(summary_file, dtype=str, keep_default_na=False)
            # 标准化列名：去除首尾空格
            df_summary.columns = [str(col).strip() for col in df_summary.columns]
            st.info(f"📊 汇总表: {len(df_summary)} 行, {len(df_summary.columns)} 列")
            with st.expander("👀 查看汇总表数据预览"):
                st.dataframe(df_summary.head(20), use_container_width=True)
                if len(df_summary) > 20:
                    st.caption(f"仅显示前 20 行，共 {len(df_summary)} 行")
        except Exception as e:
            st.error(f"❌ 读取汇总表失败: {e}")

    if detail_file:
        try:
            df_detail_raw = pd.read_excel(detail_file, dtype=str, keep_default_na=False)
            # 标准化列名：去除首尾空格
            df_detail_raw.columns = [str(col).strip() for col in df_detail_raw.columns]
            st.info(f"📋 明细表: {len(df_detail_raw)} 行, {len(df_detail_raw.columns)} 列")
            with st.expander("👀 查看明细表原始数据预览"):
                st.dataframe(df_detail_raw.head(20), use_container_width=True)
                if len(df_detail_raw) > 20:
                    st.caption(f"仅显示前 20 行，共 {len(df_detail_raw)} 行")

            clean_option = st.checkbox(
                "🧹 清洗明细表（填充合并单元格空白）",
                value=True,
                key="clean_detail",
                help="勾选后，可指定不填充的列（如金额列），避免数值重复。"
            )

            exclude_columns = []
            if clean_option:
                default_exclude = [col for col in df_detail_raw.columns if is_numeric_column(col)]
                with st.expander("⚙️ 高级设置：选择不进行填充的列（通常为金额、数量列）", expanded=True):
                    st.markdown("以下列将被排除在填充之外，保持原样。您可手动调整。")
                    exclude_columns = st.multiselect(
                        "不填充的列",
                        options=df_detail_raw.columns.tolist(),
                        default=default_exclude,
                        key="exclude_cols"
                    )
                    st.caption(f"当前选择的排除列: {exclude_columns}")

                df_detail = clean_dataframe(df_detail_raw, exclude_columns)
                st.success("✅ 明细表清洗完成（指定列未填充）")

                # 验证排除列是否真的未填充（多行对比，并用 repr 显示精确值）
                if exclude_columns:
                    with st.expander("🔍 验证排除列是否未填充（前5行对比）"):
                        for col in exclude_columns:
                            st.markdown(f"**{col}**")
                            raw_vals = df_detail_raw[col].head(5).tolist()
                            cleaned_vals = df_detail[col].head(5).tolist()
                            # 用 repr 显示空字符串
                            raw_vals_repr = [repr(v) for v in raw_vals]
                            cleaned_vals_repr = [repr(v) for v in cleaned_vals]
                            changed = any(r != c for r, c in zip(raw_vals, cleaned_vals))
                            compare_df = pd.DataFrame({
                                "行号": [f"第{i+1}行" for i in range(5)],
                                "原始值 (repr)": raw_vals_repr,
                                "清洗后值 (repr)": cleaned_vals_repr,
                                "是否一致": [r == c for r, c in zip(raw_vals, cleaned_vals)]
                            })
                            st.dataframe(compare_df, use_container_width=True)
                            if changed:
                                st.error(f"❌ 列 '{col}' 的值发生了变化！请检查是否应该被排除，或列名是否完全匹配（含括号、空格）。")
                            else:
                                st.success(f"✅ 列 '{col}' 所有值未变化。")
                else:
                    st.info("未选择任何排除列，所有列都会被填充。")

                # 订单号列对比
                detail_recommended, _ = detect_column(df_detail_raw, "明细表原始")
                if detail_recommended:
                    with st.expander("🔍 清洗前后订单号列对比（前10行）"):
                        raw_sample = df_detail_raw[detail_recommended].head(10).tolist()
                        cleaned_sample = df_detail[detail_recommended].head(10).tolist()
                        compare_df = pd.DataFrame({
                            "清洗前订单号": raw_sample,
                            "清洗后订单号": cleaned_sample,
                            "是否一致": [r == c for r, c in zip(raw_sample, cleaned_sample)]
                        })
                        st.dataframe(compare_df, use_container_width=True)

                with st.expander("👀 查看清洗后的明细表预览"):
                    st.dataframe(df_detail.head(20), use_container_width=True)

                # 下载清洗后明细表
                df_download = df_detail.copy()
                if detail_recommended:
                    df_download[detail_recommended] = df_download[detail_recommended].apply(safe_order_str)

                output_detail = io.BytesIO()
                with pd.ExcelWriter(output_detail, engine='xlsxwriter') as writer:
                    df_download.to_excel(writer, index=False, sheet_name='清洗后明细')
                output_detail.seek(0)

                st.download_button(
                    label="📥 下载清洗后的明细表 (Excel)",
                    data=output_detail,
                    file_name=f"清洗后明细表_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            else:
                df_detail = df_detail_raw
        except Exception as e:
            st.error(f"❌ 读取明细表失败: {e}")

    st.markdown("---")

    if df_summary is not None and df_detail is not None:
        st.subheader("🏷️ 选择订单号列")

        summary_recommended, summary_sorted = detect_column(df_summary, "汇总表")
        detail_recommended, detail_sorted = detect_column(df_detail, "明细表")

        col1, col2 = st.columns(2)
        with col1:
            st.markdown("**汇总表订单号列：**")
            if summary_recommended:
                st.caption(f"💡 推荐列: **{summary_recommended}**")
            summary_col = st.selectbox(
                "选择汇总表中的订单号列",
                options=summary_sorted,
                index=summary_sorted.index(summary_recommended) if summary_recommended in summary_sorted else 0,
                key="summary_col_select"
            )
        with col2:
            st.markdown("**明细表订单号列：**")
            if detail_recommended:
                st.caption(f"💡 推荐列: **{detail_recommended}**")
            detail_col = st.selectbox(
                "选择明细表中的订单号列",
                options=detail_sorted,
                index=detail_sorted.index(detail_recommended) if detail_recommended in detail_sorted else 0,
                key="detail_col_select"
            )

        st.markdown("---")

        if summary_col and detail_col:
            st.info(f"📌 将使用以下列进行匹配：\n\n"
                    f"- 汇总表：**{summary_col}**\n"
                    f"- 明细表：**{detail_col}**")

        if st.button("🔗 开始匹配", type="primary", use_container_width=True):
            if not summary_col or not detail_col:
                st.error("❌ 请先选择订单号列！")
            else:
                with st.spinner("正在匹配数据..."):
                    try:
                        df_summary[summary_col] = df_summary[summary_col].apply(safe_order_str)
                        df_detail[detail_col] = df_detail[detail_col].apply(safe_order_str)

                        with st.expander("🔍 查看清理后的订单号样例（前10条）"):
                            st.write("汇总表订单号样例：", df_summary[summary_col].head(10).tolist())
                            st.write("明细表订单号样例：", df_detail[detail_col].head(10).tolist())

                        summary_orders = df_summary[summary_col][df_summary[summary_col] != ''].unique()
                        order_set = set(summary_orders)

                        matched = df_detail[df_detail[detail_col].isin(order_set)]

                        if matched.empty:
                            st.warning("⚠️ 没有找到任何匹配的订单！")
                            st.info("💡 您可下载清洗后的明细表，并检查上方转换后的订单号样例，确认格式是否一致。")
                        else:
                            matched = convert_numeric_columns(matched)

                            st.success(f"✅ 匹配完成！找到 **{len(matched)}** 条匹配记录")

                            col1, col2, col3 = st.columns(3)
                            with col1:
                                st.metric("汇总表订单数（非空）", len(summary_orders))
                            with col2:
                                st.metric("明细表记录数", len(df_detail))
                            with col3:
                                st.metric("匹配记录数", len(matched))

                            st.subheader("📊 匹配结果预览")
                            st.dataframe(matched.head(20), use_container_width=True)
                            if len(matched) > 20:
                                st.caption(f"仅显示前 20 条记录，共 {len(matched)} 条")

                            output = io.BytesIO()
                            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                                matched.to_excel(writer, index=False, sheet_name='匹配结果')
                                workbook = writer.book
                                worksheet = writer.sheets['匹配结果']
                                col_idx = matched.columns.get_loc(detail_col)
                                col_letter = col_num_to_letter(col_idx)
                                text_format = workbook.add_format({'num_format': '@'})
                                worksheet.set_column(f'{col_letter}:{col_letter}', None, text_format)
                            output.seek(0)

                            st.markdown("---")
                            st.subheader("📥 下载结果")
                            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                            st.download_button(
                                label="📦 下载匹配结果 (Excel)",
                                data=output,
                                file_name=f"匹配结果_{timestamp}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True
                            )
                            st.info("💡 订单号列已设置为文本格式，金额列、数量列已自动转换为数值，无需手动修改。")
                    except Exception as e:
                        st.error(f"❌ 匹配过程中出现错误: {e}")
                        with st.expander("查看详细错误信息"):
                            import traceback
                            st.code(traceback.format_exc())


if __name__ == "__main__":
    main()