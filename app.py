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
    """自动识别订单号列，返回(推荐列名, 排序后列名列表)"""
    keywords = [
        "订单号", "订单编号", "订单", "编号","关联订单号","关联业务单号"
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


def is_amount_column(col_name):
    """根据列名判断是否为金额类列（不进行填充，但导出时转为数值）"""
    amount_keywords = [
        "总价", "金额", "单价","数量", "运费", "改价", "实付款", "结算价",
        "货品总价", "数量(Quantity)", "单价(Unit Price)", "订单总价(Order Amount)",
        "运费(Shipping Fee)", "预付款(Initial Payment)", "尾款(Balance Payment)",
        "price", "amount", "freight", "payment", "settlement","Order Amount"
    ]
    col_lower = str(col_name).lower()
    for kw in amount_keywords:
        if kw.lower() in col_lower:
            return True
    return False


def clean_dataframe(df):
    """
    清洗DataFrame：对非金额类列进行向前填充，金额类列保持不变。
    先将空字符串替换为 NaN，填充后再将 NaN 替换回空字符串。
    """
    # 将空字符串替换为 NaN
    df_clean = df.replace(r'^\s*$', np.nan, regex=True)
    for col in df_clean.columns:
        if not is_amount_column(col):
            # 对非金额列进行填充
            df_clean[col] = df_clean[col].ffill()
    # 将 NaN 替换回空字符串
    df_clean = df_clean.fillna('')
    return df_clean


def safe_order_str(x):
    """对订单号字符串进行简单清理：去除首尾空格。"""
    if pd.isna(x):
        return ''
    return str(x).strip()


def convert_amount_columns_to_numeric(df):
    """
    将DataFrame中的金额类列转换为数值类型（float），
    无法转换的变为 NaN。
    """
    df_numeric = df.copy()
    for col in df_numeric.columns:
        if is_amount_column(col):
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
        ✅ 明细表智能清洗：仅对非金额列（如订单号）填充合并单元格空白，金额列保持原样  
        ✅ 可下载清洗后明细表（订单号列自动转文本）  
        ✅ 清洗前后订单号对比，便于核对  
        ✅ 匹配结果中金额列自动转换为数值，无需手动修改格式  
        """)
        st.markdown("---")
        st.markdown("### 版本信息")
        st.info("版本: v1.10.0（金额列自动转数值）")

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
    df_detail_raw = None  # 保存原始明细表，用于对比
    summary_col = None
    detail_col = None

    if summary_file:
        try:
            # 以字符串形式读取，避免长数字截断
            df_summary = pd.read_excel(summary_file, dtype=str, keep_default_na=False)
            st.info(f"📊 汇总表: {len(df_summary)} 行, {len(df_summary.columns)} 列")
        except Exception as e:
            st.error(f"❌ 读取汇总表失败: {e}")

    if detail_file:
        try:
            # 保存原始数据（用于对比）
            df_detail_raw = pd.read_excel(detail_file, dtype=str, keep_default_na=False)
            st.info(f"📋 明细表: {len(df_detail_raw)} 行, {len(df_detail_raw.columns)} 列")

            clean_option = st.checkbox(
                "🧹 清洗明细表（仅对非金额列填充合并单元格空白，金额列保持不变）",
                value=True,
                key="clean_detail",
                help="勾选后，将对非金额列（如订单号）进行向下填充，金额列（如总价、实付款）保持原样。"
            )
            if clean_option:
                df_detail = clean_dataframe(df_detail_raw)
                st.success("✅ 明细表清洗完成（非金额列已填充）")

                # 清洗前后订单号列对比（如果检测到推荐列）
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

                # 预览清洗后的明细表
                with st.expander("👀 查看清洗后的明细表预览"):
                    st.dataframe(df_detail.head(20), use_container_width=True)

                # 生成下载清洗后明细表（对推荐列应用 safe_order_str 去除空格）
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
                df_detail = df_detail_raw  # 不清洗则使用原始数据
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
                        # 对订单号列进行简单清理（去除空格）
                        df_summary[summary_col] = df_summary[summary_col].apply(safe_order_str)
                        df_detail[detail_col] = df_detail[detail_col].apply(safe_order_str)

                        # 调试：显示转换后的订单号样例
                        with st.expander("🔍 查看清理后的订单号样例（前10条）"):
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
                            # 将金额列转换为数值类型（便于后续计算和显示）
                            matched = convert_amount_columns_to_numeric(matched)

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

                            # 生成下载文件，并对订单号列设置文本格式，金额列自动为数值
                            output = io.BytesIO()
                            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                                matched.to_excel(writer, index=False, sheet_name='匹配结果')
                                workbook = writer.book
                                worksheet = writer.sheets['匹配结果']
                                # 将订单号列设为文本格式
                                col_idx = matched.columns.get_loc(detail_col)
                                col_letter = col_num_to_letter(col_idx)
                                text_format = workbook.add_format({'num_format': '@'})
                                worksheet.set_column(f'{col_letter}:{col_letter}', None, text_format)
                                # 可选：将金额列设为数值格式（但默认就是数值，无需额外设置）
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
                            st.info("💡 订单号列已设置为文本格式，金额列已自动转换为数值，无需手动修改。")
                    except Exception as e:
                        st.error(f"❌ 匹配过程中出现错误: {e}")
                        with st.expander("查看详细错误信息"):
                            import traceback
                            st.code(traceback.format_exc())


if __name__ == "__main__":
    main()