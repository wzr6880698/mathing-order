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
    """
    # 先处理空列名
    cleaned_columns = []
    for i, col in enumerate(df.columns):
        if pd.isna(col) or str(col).strip() == '':
            cleaned_columns.append(f"列{i+1}")
        else:
            cleaned_columns.append(str(col))
    df.columns = cleaned_columns

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


def is_numeric_column(col_name):
    """判断某列是否为数字相关列（金额、数量等），用于清洗时排除填充"""
    numeric_keywords = [
        "总价", "金额", "单价", "运费", "改价", "实付款", "结算价", "货品总价",
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


def is_total_column(col_name):
    """判断某列是否为合计类列（需还原合并单元格效果）"""
    total_keywords = [
        "总价", "运费", "改价", "实付款", "结算价", "货品总价",
        "order amount", "shipping fee", "discount amount",
        "total amount", "fee", "discount", "shipping",
        "initial payment", "balance payment", "tax"
    ]
    col_lower = str(col_name).lower()
    for kw in total_keywords:
        if kw.lower() in col_lower:
            return True
    return False


def clean_dataframe(df, exclude_columns=None):
    """
    清洗DataFrame：对非排除列进行向前填充，排除列保持原始值不变。
    """
    if exclude_columns is None:
        exclude_columns = []
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


def convert_numeric_columns(df):
    """将数字相关列转换为数值类型（float），无法转换的变为 NaN。"""
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
        ✅ 明细表自动清洗（填充非金额列空白）  
        ✅ 自动还原合计列（如订单总价）的合并单元格效果  
        ✅ 单价、数量等明细列保持不变  
        ✅ 匹配结果中金额列自动转为数值  
        """)
        st.markdown("---")
        st.markdown("### 版本信息")
        st.info("版本: v3.1.0（智能还原合计列）")

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
            df_summary = pd.read_excel(summary_file, dtype=str, keep_default_na=False)
            df_summary.columns = [str(col).strip() for col in df_summary.columns]
            st.info(f"📊 汇总表: {len(df_summary)} 行, {len(df_summary.columns)} 列")
            with st.expander("👀 查看汇总表数据预览"):
                st.dataframe(df_summary.head(20), use_container_width=True)
        except Exception as e:
            st.error(f"❌ 读取汇总表失败: {e}")

    if detail_file:
        try:
            df_detail_raw = pd.read_excel(detail_file, dtype=str, keep_default_na=False)
            df_detail_raw.columns = [str(col).strip() for col in df_detail_raw.columns]
            st.info(f"📋 明细表: {len(df_detail_raw)} 行, {len(df_detail_raw.columns)} 列")
            with st.expander("👀 查看明细表原始数据预览（pandas已填充合并单元格）"):
                st.dataframe(df_detail_raw.head(20), use_container_width=True)

            # 清洗明细表（排除所有数字列，避免二次填充）
            exclude_columns = [col for col in df_detail_raw.columns if is_numeric_column(col)]
            df_detail = clean_dataframe(df_detail_raw, exclude_columns)
            st.success("✅ 明细表清洗完成（数字列已排除填充）")

            # 显示清洗后的明细表（此时数字列仍为pandas填充状态）
            with st.expander("👀 查看清洗后的明细表（数字列仍为填充状态）"):
                st.dataframe(df_detail.head(20), use_container_width=True)

        except Exception as e:
            st.error(f"❌ 读取明细表失败: {e}")

    st.markdown("---")

    if df_summary is not None and df_detail is not None:
        st.subheader("🏷️ 选择订单号列")

        # 检测列
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
                        # 清理订单号列（去除空格）
                        df_summary[summary_col] = df_summary[summary_col].apply(safe_order_str)
                        df_detail[detail_col] = df_detail[detail_col].apply(safe_order_str)

                        # 自动还原合计列（仅保留每组订单的第一行）
                        total_cols = [col for col in df_detail.columns if is_total_column(col)]
                        if total_cols:
                            # 按订单号分组，每组第一行保留，其余清空
                            mask = df_detail.duplicated(subset=[detail_col], keep='first')
                            for col in total_cols:
                                df_detail.loc[mask, col] = ''
                            st.info(f"🔄 已自动还原合计列: {total_cols}")

                        # 显示最终处理后的明细表预览
                        with st.expander("👀 查看最终处理后的明细表（合计列已还原）"):
                            st.dataframe(df_detail.head(20), use_container_width=True)

                        # 汇总表有效订单集合
                        summary_orders = df_summary[summary_col][df_summary[summary_col] != ''].unique()
                        order_set = set(summary_orders)

                        # 匹配
                        matched = df_detail[df_detail[detail_col].isin(order_set)]

                        if matched.empty:
                            st.warning("⚠️ 没有找到任何匹配的订单！")
                            st.info("💡 您可检查上方订单号样例，确认格式是否一致。")
                        else:
                            # 将数字列转换为数值
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

                            # 生成下载文件
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
                            st.info("💡 订单号列已设为文本格式，合计列已自动还原合并单元格效果。")
                    except Exception as e:
                        st.error(f"❌ 匹配过程中出现错误: {e}")
                        with st.expander("查看详细错误信息"):
                            import traceback
                            st.code(traceback.format_exc())


if __name__ == "__main__":
    main()