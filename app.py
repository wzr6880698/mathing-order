import pandas as pd
import streamlit as st
import re
import io
from datetime import datetime
import openpyxl


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
    # 包含中文字符的通常不是订单号（内销订单号可能包含特定前缀，保留此规则）
    if re.search('[\u4e00-\u9fff]', s):
        return False
    # 允许字母、数字、短横线、下划线、斜杠、点
    if re.match(r'^[A-Za-z0-9\-_/\.]+$', s):
        return True
    return False


def fix_excel_column_names(file_obj):
    """
    从Excel文件中读取真实列名（解决Unnamed问题）
    返回：修复后的DataFrame + 原始列名列表
    """
    # 第一步：用openpyxl读取原始列名（保留合并单元格/真实表头）
    wb = openpyxl.load_workbook(file_obj, data_only=True)
    ws = wb.active

    # 获取表头行（默认第一行，可适配多行表头）
    header_row = []
    for cell in ws[1]:  # 读取第一行作为表头
        cell_value = cell.value
        if cell_value is None:
            # 空单元格用"列字母+行号"命名（如A1、B1），友好且唯一
            header_row.append(f"{cell.column_letter}{cell.row}")
        else:
            # 清理表头中的特殊字符和空格
            clean_val = str(cell_value).strip().replace('\n', ' ').replace('\t', ' ')
            # 去重处理
            if clean_val in header_row:
                count = header_row.count(clean_val) + 1
                header_row.append(f"{clean_val}_{count}")
            else:
                header_row.append(clean_val)

    # 第二步：读取数据（跳过表头行）
    df = pd.read_excel(file_obj, header=None, skiprows=1)

    # 第三步：设置修复后的列名
    df.columns = header_row[:len(df.columns)]  # 防止列数不匹配

    # 第四步：处理空行和全空列
    df = df.dropna(how='all')  # 删除全空行
    df = df.dropna(axis=1, how='all')  # 删除全空列

    return df


def fill_merged_cells(df, order_col):
    """填充合并单元格导致的空值（针对订单号列）"""
    # 向前填充订单号列的空值（合并单元格在Excel中读取后为空，需继承上一行订单号）
    if order_col in df.columns:
        # 先将NaN转为空字符串，再填充
        df[order_col] = df[order_col].fillna('').astype(str)
        df[order_col] = df[order_col].replace('', method='ffill')
    return df


def detect_domestic_order_column(df):
    """
    智能识别内销订单号列（优先级排序）
    返回：最可能的订单号列名 / None
    """
    # 内销订单号列关键词（按优先级排序）
    domestic_keywords = [
        # 高优先级：明确的订单号关键词
        '订单号', '内销订单号', '订单编号', '内销编号', '单号', '内销单号',
        # 中优先级：通用编号关键词
        '编号', '序号', 'ID', 'id', 'No', 'NO', 'no',
        # 低优先级：首列（内销明细订单号通常在首列）
        df.columns[0] if len(df.columns) > 0 else None
    ]

    # 遍历关键词找匹配列
    for keyword in domestic_keywords:
        if keyword is None:
            continue
        # 模糊匹配列名
        for col in df.columns:
            if keyword in str(col):
                return col

    # 如果没有匹配，返回非空值最多的列（订单号列通常非空值多）
    non_null_counts = df.notna().sum()
    if not non_null_counts.empty:
        return non_null_counts.idxmax()

    return None


def detect_export_order_column(df):
    """智能识别外销订单号列"""
    export_keywords = [
        '订单号', '订单编号', 'Order Number', 'order number',
        'orderno', 'Order No', 'order no', '订单ID'
    ]

    for keyword in export_keywords:
        for col in df.columns:
            if keyword.lower() in str(col).lower():
                return col

    non_null_counts = df.notna().sum()
    if not non_null_counts.empty:
        return non_null_counts.idxmax()

    return None


def detect_file_type_and_order_col(df, file_name):
    """
    综合识别：文件类型 + 推荐订单号列
    返回：(file_type: 'domestic'/'export', recommend_col: str)
    """
    # 先通过文件名判断
    if any(keyword in file_name.lower() for keyword in ['内销', 'domestic']):
        return 'domestic', detect_domestic_order_column(df)
    elif any(keyword in file_name.lower() for keyword in ['外销', 'export']):
        return 'export', detect_export_order_column(df)

    # 文件名无标识，通过内容判断
    # 统计内销/外销特征词
    col_names = ' '.join([str(col).lower() for col in df.columns])
    domestic_score = sum(1 for kw in ['内销', '国内'] if kw in col_names)
    export_score = sum(1 for kw in ['外销', '出口', 'export'] if kw in col_names)

    if domestic_score > export_score:
        return 'domestic', detect_domestic_order_column(df)
    else:
        return 'export', detect_export_order_column(df)


def main():
    st.set_page_config(
        page_title="订单匹配工具（内销+外销通用）",
        page_icon="🔗",
        layout="wide",
        initial_sidebar_state="expanded"
    )

    # 侧边栏说明
    with st.sidebar:
        st.header("📖 使用说明")
        st.markdown("""
        ### 核心优势
        ✅ 智能识别内销/外销文件，无需手动切换
        ✅ 保留原始列名，解决Unnamed列名显示问题
        ✅ 自动填充内销明细表合并单元格空值
        ✅ 精准推荐订单号列，减少手动选择
        ✅ 订单号文本格式，避免科学计数法

        ### 操作步骤
        1. 上传汇总表（内销/外销Excel）
        2. 上传明细表（内销/外销Excel）
        3. 确认/选择订单号列（系统已智能推荐）
        4. 点击匹配，查看结果
        5. 下载完整匹配结果
        """)
        st.markdown("---")
        st.info("版本: v3.0.0（智能列名+最优用户体验）")

    # 主界面
    st.title("🔗 订单匹配工具（内销+外销通用）")
    st.markdown("### 智能识别列名 | 完美兼容内销/外销格式 | 友好用户体验")
    st.markdown("---")

    # 文件上传区域
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("📊 汇总表")
        summary_file = st.file_uploader(
            "上传汇总Excel文件（内销/外销均可）",
            type=['xlsx', 'xls'],
            key="summary_file",
            help="支持.xlsx/.xls格式，自动识别列名"
        )
        if summary_file:
            st.success(f"✅ 已上传: {summary_file.name}")

    with col2:
        st.subheader("📋 明细表")
        detail_file = st.file_uploader(
            "上传明细Excel文件（内销/外销均可）",
            type=['xlsx', 'xls'],
            key="detail_file",
            help="内销文件自动处理合并单元格空值"
        )
        if detail_file:
            st.success(f"✅ 已上传: {detail_file.name}")

    # 数据读取和预处理（核心：修复列名）
    df_summary = None
    df_detail = None
    summary_type = None
    detail_type = None
    summary_recommend_col = None
    detail_recommend_col = None

    # 读取汇总表
    if summary_file:
        try:
            # 关键：使用修复列名的函数读取
            df_summary = fix_excel_column_names(summary_file)
            st.info(f"📊 汇总表：{len(df_summary)} 行 | {len(df_summary.columns)} 列")

            # 识别文件类型和推荐列
            summary_type, summary_recommend_col = detect_file_type_and_order_col(
                df_summary, summary_file.name
            )
            type_text = "内销" if summary_type == "domestic" else "外销"
            st.success(f"📌 自动识别：{type_text}格式")
            if summary_recommend_col:
                st.info(f"💡 推荐订单号列：**{summary_recommend_col}**")

        except Exception as e:
            st.error(f"❌ 读取汇总表失败：{str(e)}")
            st.warning("建议检查文件格式，确保是标准Excel文件")

    # 读取明细表
    if detail_file:
        try:
            # 关键：使用修复列名的函数读取
            df_detail = fix_excel_column_names(detail_file)
            st.info(f"📋 明细表：{len(df_detail)} 行 | {len(df_detail.columns)} 列")

            # 识别文件类型和推荐列
            detail_type, detail_recommend_col = detect_file_type_and_order_col(
                df_detail, detail_file.name
            )
            type_text = "内销" if detail_type == "domestic" else "外销"
            st.success(f"📌 自动识别：{type_text}格式")

            # 内销明细表自动填充合并单元格空值
            if detail_type == "domestic" and detail_recommend_col:
                df_detail = fill_merged_cells(df_detail, detail_recommend_col)
                st.success(f"✅ 已填充内销明细表【{detail_recommend_col}】列合并单元格空值")

            if detail_recommend_col:
                st.info(f"💡 推荐订单号列：**{detail_recommend_col}**")

        except Exception as e:
            st.error(f"❌ 读取明细表失败：{str(e)}")
            st.warning("建议检查文件格式，确保是标准Excel文件")

    st.markdown("---")

    # 列选择区域（友好列名展示）
    if df_summary is not None and df_detail is not None:
        st.subheader("🏷️ 选择订单号列（系统已智能推荐）")

        col1, col2 = st.columns(2)
        with col1:
            st.markdown(f"**汇总表（{summary_type}格式）：**")
            # 列选择框：显示真实友好的列名
            summary_col = st.selectbox(
                "选择汇总表订单号列",
                options=df_summary.columns.tolist(),
                index=df_summary.columns.tolist().index(summary_recommend_col)
                if summary_recommend_col in df_summary.columns else 0,
                key="summary_col",
                help="选择包含订单号的列，系统已推荐最优列"
            )

        with col2:
            st.markdown(f"**明细表（{detail_type}格式）：**")
            # 列选择框：显示真实友好的列名
            detail_col = st.selectbox(
                "选择明细表订单号列",
                options=df_detail.columns.tolist(),
                index=df_detail.columns.tolist().index(detail_recommend_col)
                if detail_recommend_col in df_detail.columns else 0,
                key="detail_col",
                help="内销文件已自动填充合并单元格空值"
            )

        # 手动填充按钮（内销专用）
        if detail_type == "domestic":
            st.markdown("---")
            if st.button("🔄 重新填充明细表订单号列空值", use_container_width=True):
                df_detail = fill_merged_cells(df_detail, detail_col)
                st.success(f"✅ 已重新填充【{detail_col}】列的空值（合并单元格）")

        # 显示列选择确认信息
        st.markdown("---")
        st.info(f"""
        📌 匹配配置确认：
        - 汇总表订单号列：**{summary_col}**
        - 明细表订单号列：**{detail_col}**
        - 文件类型：汇总表({summary_type}) | 明细表({detail_type})
        """)

        # 匹配按钮
        if st.button("🔗 开始智能匹配", type="primary", use_container_width=True):
            with st.spinner("正在执行订单匹配，请稍候..."):
                try:
                    # 数据预处理
                    # 1. 确保订单号列为字符串格式
                    df_summary[summary_col] = df_summary[summary_col].astype(str).str.strip()
                    df_detail[detail_col] = df_detail[detail_col].astype(str).str.strip()

                    # 2. 过滤无效订单号（空字符串/纯空格）
                    valid_summary_orders = df_summary[
                        df_summary[summary_col] != ''
                        ][summary_col].unique()
                    order_set = set(valid_summary_orders)

                    # 3. 执行匹配
                    matched_df = df_detail[df_detail[detail_col].isin(order_set)]

                    # 结果展示
                    if len(matched_df) == 0:
                        st.warning("⚠️ 未找到匹配的订单记录！")
                        with st.expander("📈 匹配分析（帮助排查问题）"):
                            st.markdown(f"- 汇总表有效订单数：{len(valid_summary_orders)}")
                            st.markdown(f"- 明细表有效记录数：{len(df_detail[df_detail[detail_col] != ''])}")
                            st.markdown(
                                f"- 汇总表订单号样本：{list(valid_summary_orders[:5]) if len(valid_summary_orders) > 0 else '无'}")
                            st.markdown(
                                f"- 明细表订单号样本：{list(df_detail[df_detail[detail_col] != ''][detail_col].unique()[:5]) if len(df_detail) > 0 else '无'}")
                    else:
                        st.success(f"✅ 匹配完成！共找到 **{len(matched_df)}** 条匹配记录")

                        # 统计信息
                        stat_col1, stat_col2, stat_col3, stat_col4 = st.columns(4)
                        with stat_col1:
                            st.metric("汇总表订单数", len(valid_summary_orders))
                        with stat_col2:
                            st.metric("明细表总记录数", len(df_detail))
                        with stat_col3:
                            st.metric("匹配记录数", len(matched_df))
                        with stat_col4:
                            valid_detail = len(df_detail[df_detail[detail_col] != ''])
                            match_rate = (len(matched_df) / valid_detail) * 100 if valid_detail > 0 else 0
                            st.metric("匹配率", f"{match_rate:.1f}%")

                        # 结果预览
                        st.subheader("📊 匹配结果预览")
                        st.dataframe(matched_df.head(30), use_container_width=True)
                        if len(matched_df) > 30:
                            st.caption(f"仅展示前30条，共{len(matched_df)}条匹配记录")

                        # 生成下载文件
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                            matched_df.to_excel(writer, index=False, sheet_name='匹配结果')
                            # 设置订单号列为文本格式
                            workbook = writer.book
                            worksheet = writer.sheets['匹配结果']
                            col_idx = matched_df.columns.get_loc(detail_col)
                            col_letter = col_num_to_letter(col_idx)
                            text_format = workbook.add_format({'num_format': '@'})
                            worksheet.set_column(f'{col_letter}:{col_letter}', None, text_format)

                        output.seek(0)

                        # 下载按钮
                        st.markdown("---")
                        st.subheader("📥 下载匹配结果")
                        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                        st.download_button(
                            label="📦 下载完整匹配结果 (Excel)",
                            data=output,
                            file_name=f"订单匹配结果_{timestamp}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                        st.info("💡 订单号列已设置为文本格式，避免科学计数法；内销文件合并单元格空值已填充")

                except Exception as e:
                    st.error(f"❌ 匹配过程出错：{str(e)}")
                    with st.expander("查看详细错误信息"):
                        import traceback
                        st.code(traceback.format_exc())


if __name__ == "__main__":
    main()