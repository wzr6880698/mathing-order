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
    # 包含中文字符的通常不是订单号
    if re.search('[\u4e00-\u9fff]', s):
        return False
    # 允许字母、数字、短横线、下划线、斜杠、点
    if re.match(r'^[A-Za-z0-9\-_/\.]+$', s):
        return True
    return False


def detect_column(df, sheet_desc):
    """
    自动识别订单号列，结合列名关键词和列内数据特征。
    返回 (推荐的列名, 所有列名按得分排序)
    """
    # 关键词列表
    keywords = [
        "订单号", "订单编号", "订单", "编号",
        "order number", "order no", "orderno", "order id",
        "order", "id", "编号"
    ]

    # 1. 关键词得分
    kw_scores = {}
    for col in df.columns:
        col_lower = str(col).lower()
        score = 0
        for kw in keywords:
            if kw.lower() in col_lower:
                score += 1
        kw_scores[col] = score

    # 2. 内容特征得分
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

    # 3. 综合得分
    total_scores = {}
    for col in df.columns:
        kw = kw_scores[col]
        content = content_scores[col]
        if kw > 0:
            total = kw * 10 + content
        else:
            total = content
        total_scores[col] = total

    # 4. 按得分排序
    sorted_columns = sorted(df.columns.tolist(), key=lambda c: total_scores[c], reverse=True)

    # 找出最高分的列
    max_score = max(total_scores.values()) if total_scores else 0
    recommended = sorted_columns[0] if sorted_columns and max_score > 0 else None

    return recommended, sorted_columns


def clean_dataframe(df):
    """
    清洗DataFrame：对所有列执行向前填充（ffill），
    以填充因合并单元格导致的空白单元格。
    """
    return df.ffill()


def main():
    st.set_page_config(
        page_title="订单匹配工具",
        page_icon="🔗",
        layout="wide",
        initial_sidebar_state="expanded"
    )

    # 侧边栏说明
    with st.sidebar:
        st.header("📖 使用说明")
        st.markdown("""
        ### 操作步骤：

        1. **上传汇总表** 📊
           - 包含订单汇总信息的Excel文件

        2. **上传明细表** 📋
           - 包含订单明细信息的Excel文件

        3. **选择订单号列** 🏷️
           - 系统会自动识别推荐
           - 如不准确可手动选择

        4. **开始匹配** 🔗
           - 点击按钮执行匹配

        5. **下载结果** ⬇️
           - 下载匹配后的Excel文件

        ---

        ### 功能说明：

        ✅ 自动识别订单号列
        ✅ 支持手动选择列
        ✅ 明细表智能清洗（填充所有合并单元格空白）
        ✅ 订单号自动转为文本格式，避免科学计数法
        ✅ 匹配结果中订单号无`nan`显示
        """)

        st.markdown("---")
        st.markdown("### 版本信息")
        st.info("版本: v1.2.0（全面清洗+数值列保留）")

    # 主界面
    st.title("🔗 订单匹配工具")
    st.markdown("根据订单号匹配汇总表和明细表数据")
    st.markdown("---")

    # 创建两列布局
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("📊 汇总表")
        summary_file = st.file_uploader(
            "上传汇总信息 Excel 文件",
            type=['xlsx', 'xls'],
            key="summary_file"
        )

        if summary_file:
            st.success(f"✅ 已上传: {summary_file.name}")

    with col2:
        st.subheader("📋 明细表")
        detail_file = st.file_uploader(
            "上传明细信息 Excel 文件",
            type=['xlsx', 'xls'],
            key="detail_file"
        )

        if detail_file:
            st.success(f"✅ 已上传: {detail_file.name}")

    # 读取数据
    df_summary = None
    df_detail = None
    summary_col = None
    detail_col = None
    summary_columns = []
    detail_columns = []

    if summary_file:
        try:
            df_summary = pd.read_excel(summary_file)
            st.info(f"📊 汇总表: {len(df_summary)} 行, {len(df_summary.columns)} 列")
        except Exception as e:
            st.error(f"❌ 读取汇总表失败: {e}")

    if detail_file:
        try:
            df_detail = pd.read_excel(detail_file)
            st.info(f"📋 明细表: {len(df_detail)} 行, {len(df_detail.columns)} 列")

            # 明细表清洗选项
            clean_option = st.checkbox(
                "🧹 清洗明细表（自动填充所有因合并单元格导致的空白）",
                value=True,
                key="clean_detail",
                help="勾选后，将对所有列进行向下填充，确保订单号、金额等信息完整。"
            )
            if clean_option:
                df_detail = clean_dataframe(df_detail)
                st.success("✅ 明细表清洗完成（所有空白单元格已填充）")
        except Exception as e:
            st.error(f"❌ 读取明细表失败: {e}")

    st.markdown("---")

    # 选择订单号列
    if df_summary is not None and df_detail is not None:
        st.subheader("🏷️ 选择订单号列")

        # 自动识别
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

        # 确认信息
        if summary_col and detail_col:
            st.info(f"📌 将使用以下列进行匹配：\n\n"
                    f"- 汇总表：**{summary_col}**\n"
                    f"- 明细表：**{detail_col}**")

        # 开始匹配按钮
        if st.button("🔗 开始匹配", type="primary", use_container_width=True):
            if not summary_col or not detail_col:
                st.error("❌ 请先选择订单号列！")
            else:
                with st.spinner("正在匹配数据..."):
                    try:
                        # 单独处理订单号列：将 NaN 替换为空字符串，再转为字符串（避免出现 "nan"）
                        df_summary[summary_col] = df_summary[summary_col].fillna('').astype(str)
                        df_detail[detail_col] = df_detail[detail_col].fillna('').astype(str)

                        # 构建汇总表订单号集合（排除空字符串）
                        summary_orders = df_summary[summary_col][df_summary[summary_col] != ''].unique()
                        order_set = set(summary_orders)

                        # 匹配（明细表订单号在集合中）
                        matched = df_detail[df_detail[detail_col].isin(order_set)]

                        if matched.empty:
                            st.warning("⚠️ 没有找到任何匹配的订单！")
                        else:
                            # 显示匹配结果
                            st.success(f"✅ 匹配完成！找到 **{len(matched)}** 条匹配记录")

                            # 显示统计信息
                            col1, col2, col3 = st.columns(3)
                            with col1:
                                st.metric("汇总表订单数（非空）", len(summary_orders))
                            with col2:
                                st.metric("明细表记录数", len(df_detail))
                            with col3:
                                st.metric("匹配记录数", len(matched))

                            # 预览数据
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

                                # 获取订单号列位置
                                col_idx = matched.columns.get_loc(detail_col)
                                col_letter = col_num_to_letter(col_idx)

                                # 设置订单号列为文本格式（避免科学计数法）
                                text_format = workbook.add_format({'num_format': '@'})
                                worksheet.set_column(f'{col_letter}:{col_letter}', None, text_format)

                            output.seek(0)

                            # 下载按钮
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