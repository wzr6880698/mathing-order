import pandas as pd
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
    """自动识别订单号列，返回(推荐列名, 排序后列名列表)"""
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
        total = kw * 10 + content if kw > 0 else content
        total_scores[col] = total

    sorted_columns = sorted(df.columns.tolist(), key=lambda c: total_scores[c], reverse=True)
    max_score = max(total_scores.values()) if total_scores else 0
    recommended = sorted_columns[0] if sorted_columns and max_score > 0 else None
    return recommended, sorted_columns


def clean_dataframe(df):
    """对所有列执行向前填充，填充合并单元格导致的空白"""
    return df.ffill()


def safe_order_str(x):
    """
    将订单号值安全转换为字符串，避免科学计数法，并去除首尾空格及特殊空白字符。
    - NaN -> ''
    - 整数或可转为整数的浮点数 -> 整数格式字符串 (如 1.0 -> "1")
    - 其他数值 -> 直接转字符串 (保留小数)
    - 其他 -> 直接转字符串
    最后执行 strip() 去除各类空白字符。
    """
    if pd.isna(x):
        return ''
    if isinstance(x, (int, float)):
        if x == int(x):
            s = str(int(x))
        else:
            s = str(x)
    else:
        s = str(x)
    # 去除各种空白字符，包括普通空格、\xa0等
    return s.strip()


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
        ✅ 明细表智能清洗（填充合并单元格空白）  
        ✅ 可下载清洗后明细表（订单号推荐列自动转文本）  
        ✅ 订单号自动转为文本格式，避免科学计数法，并去除空格  
        ✅ 清洗前后订单号对比功能  
        """)
        st.markdown("---")
        st.markdown("### 版本信息")
        st.info("版本: v1.7.0（增强清洗对比与匹配调试）")

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
    summary_col = None
    detail_col = None

    if summary_file:
        try:
            df_summary = pd.read_excel(summary_file)
            st.info(f"📊 汇总表: {len(df_summary)} 行, {len(df_summary.columns)} 列")
        except Exception as e:
            st.error(f"❌ 读取汇总表失败: {e}")

    if detail_file:
        try:
            # 读取原始明细表，用于后续对比
            df_detail_raw = pd.read_excel(detail_file)
            st.info(f"📋 明细表: {len(df_detail_raw)} 行, {len(df_detail_raw.columns)} 列")

            clean_option = st.checkbox(
                "🧹 清洗明细表（自动填充所有因合并单元格导致的空白）",
                value=True,
                key="clean_detail",
                help="勾选后，将对所有列进行向下填充，确保订单号、金额等信息完整。"
            )
            if clean_option:
                df_detail = clean_dataframe(df_detail_raw)
                st.success("✅ 明细表清洗完成（所有空白单元格已填充）")

                # 检测推荐订单号列
                detail_recommended, _ = detect_column(df_detail, "明细表")
                if detail_recommended:
                    st.info(f"🔍 检测到可能的订单号列：**{detail_recommended}**")

                    # 提供清洗前后订单号对比
                    with st.expander("🔎 查看清洗前后订单号列对比（前20行）"):
                        # 获取原始列和清洗后列，都转换为字符串以便显示
                        raw_vals = df_detail_raw[detail_recommended].astype(str).tolist()[:20]
                        cleaned_vals = df_detail[detail_recommended].astype(str).tolist()[:20]
                        compare_df = pd.DataFrame({
                            "原始订单号（前20）": raw_vals,
                            "清洗后订单号（前20）": cleaned_vals
                        })
                        st.dataframe(compare_df, use_container_width=True)
                        st.caption("注意：清洗后空值会被上一行填充，这是预期行为。")

                # 预览清洗后的明细表（原始数据，可能显示科学计数）
                with st.expander("👀 查看清洗后的明细表预览（所有列前20行）"):
                    st.dataframe(df_detail.head(20), use_container_width=True)

                # 生成下载清洗后明细表（对推荐列应用 safe_order_str）
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
                # 如果不清洗，直接使用原始数据
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
                index=0 if summary_sorted else None,
                key="summary_col_select"
            )
        with col2:
            st.markdown("**明细表订单号列：**")
            if detail_recommended:
                st.caption(f"💡 推荐列: **{detail_recommended}**")
            detail_col = st.selectbox(
                "选择明细表中的订单号列",
                options=detail_sorted,
                index=0 if detail_sorted else None,
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
                        # 对订单号列进行安全转换（避免科学计数法、去除各类空格）
                        df_summary[summary_col] = df_summary[summary_col].apply(safe_order_str)
                        df_detail[detail_col] = df_detail[detail_col].apply(safe_order_str)

                        # 调试：显示转换后的订单号样例
                        with st.expander("🔍 查看转换后的订单号样例（前10条）"):
                            st.write("汇总表订单号样例：", df_summary[summary_col].head(10).tolist())
                            st.write("明细表订单号样例：", df_detail[detail_col].head(10).tolist())

                        # 汇总表有效订单集合（排除空字符串）
                        summary_orders = df_summary[summary_col][df_summary[summary_col] != ''].unique()
                        order_set = set(summary_orders)

                        # 匹配
                        matched = df_detail[df_detail[detail_col].isin(order_set)]

                        if matched.empty:
                            st.warning("⚠️ 没有找到任何匹配的订单！")
                            st.info("💡 您可下载清洗后的明细表，并检查上方转换后的订单号样例，确认格式是否一致。")
                        else:
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

                            # 生成下载文件，并对订单号列设置文本格式
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
                            st.info("💡 订单号列已设置为文本格式，不会显示为科学计数法。")
                    except Exception as e:
                        st.error(f"❌ 匹配过程中出现错误: {e}")
                        with st.expander("查看详细错误信息"):
                            import traceback
                            st.code(traceback.format_exc())


if __name__ == "__main__":
    main()