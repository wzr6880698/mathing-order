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
    # 包含中文字符的通常不是订单号（内销订单号可能包含特定前缀，保留此规则）
    if re.search('[\u4e00-\u9fff]', s):
        return False
    # 允许字母、数字、短横线、下划线、斜杠、点
    if re.match(r'^[A-Za-z0-9\-_/\.]+$', s):
        return True
    return False

def detect_file_type(df, file_name):
    """
    自动识别文件类型（内销/外销）
    返回: 'domestic' (内销) / 'export' (外销) / None (无法识别)
    """
    # 外销文件特征：包含特定列名
    export_key_cols = {'订单号', '订单编号', 'Order Number', 'order number', 'orderno'}
    # 内销文件特征：假设包含内销特有列名（可根据实际内销格式调整）
    domestic_key_cols = {'内销订单号', '内销编号', 'domestic order', 'domestic no', '内销单号'}
    
    df_cols = set(str(col).lower() for col in df.columns)
    export_match = sum(1 for col in export_key_cols if col.lower() in df_cols)
    domestic_match = sum(1 for col in domestic_key_cols if col.lower() in df_cols)
    
    if export_match > 0 and domestic_match == 0:
        return 'export'
    elif domestic_match > 0 and export_match == 0:
        return 'domestic'
    else:
        # 若列名无法识别，通过订单号格式辅助判断（外销订单号通常为纯数字/字母组合，内销可能有特定前缀）
        sample_size = min(100, len(df))
        sample = df.head(sample_size)
        export_order_count = 0
        domestic_order_count = 0
        
        for col in df.columns:
            non_null = sample[col].notna().sum()
            if non_null == 0:
                continue
            values = sample[col].dropna().astype(str).tolist()
            # 外销订单号：纯数字/字母/符号组合，无中文
            export_count = sum(1 for v in values if like_order_string(v))
            # 内销订单号：假设包含特定前缀（如'D'/'DOM'等，可根据实际调整）
            domestic_count = sum(1 for v in values if re.match(r'^[DdOMom\-_0-9]+$', v) and like_order_string(v))
            
            export_order_count += export_count
            domestic_order_count += domestic_count
        
        return 'export' if export_order_count > domestic_order_count else 'domestic' if domestic_order_count > 0 else None

def detect_column(df, sheet_desc, file_type=None):
    """
    自动识别订单号列，支持内销/外销双格式
    返回 (推荐的列名, 所有列名按得分排序)
    """
    # 合并内销+外销关键词列表
    keywords = [
        # 外销关键词
        "订单号", "订单编号", "order number", "order no", "orderno", "order id", "order", "id",
        # 内销关键词（可根据实际内销格式补充调整）
        "内销订单号", "内销编号", "内销单号", "domestic order", "domestic no", "domestic id"
    ]
    
    # 1. 关键词得分（优先匹配对应文件类型的关键词）
    kw_scores = {}
    for col in df.columns:
        col_lower = str(col).lower()
        score = 0
        for kw in keywords:
            kw_lower = kw.lower()
            # 若指定了文件类型，对应类型的关键词加分权重更高
            if file_type == 'export' and kw in ["订单号", "订单编号", "order number", "order no", "orderno"]:
                if kw_lower in col_lower:
                    score += 2  # 外销关键词权重翻倍
            elif file_type == 'domestic' and kw in ["内销订单号", "内销编号", "内销单号", "domestic order"]:
                if kw_lower in col_lower:
                    score += 2  # 内销关键词权重翻倍
            else:
                if kw_lower in col_lower:
                    score += 1
        kw_scores[col] = score
    
    # 2. 内容特征得分（适配内销订单号格式）
    content_scores = {}
    sample_size = min(100, len(df))
    sample = df.head(sample_size)
    for col in df.columns:
        non_null = sample[col].notna().sum()
        if non_null == 0:
            content_scores[col] = 0.0
            continue
        values = sample[col].dropna().astype(str).tolist()
        # 内销订单号可能包含特定前缀（如'D'/'DOM'），调整匹配规则
        if file_type == 'domestic':
            like_count = sum(1 for v in values if re.match(r'^[A-Za-z0-9\-_/\.]+$', v) and (len(v) >=3 and len(v) <=50))
        else:
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

def main():
    st.set_page_config(
        page_title="订单匹配工具（内销+外销通用）",
        page_icon="🔗",
        layout="wide",
        initial_sidebar_state="expanded"
    )
    
    # 侧边栏说明（更新支持双格式）
    with st.sidebar:
        st.header("📖 使用说明")
        st.markdown("""
        ### 操作步骤：
        1. **上传汇总表** 📊
           - 支持内销/外销汇总Excel文件
        2. **上传明细表** 📋
           - 支持内销/外销明细Excel文件
        3. **选择订单号列** 🏷️
           - 系统自动识别文件类型和推荐列
           - 如不准确可手动选择
        4. **开始匹配** 🔗
           - 点击按钮执行匹配
        5. **下载结果** ⬇️
           - 下载匹配后的Excel文件
        ---
        ### 功能说明：
        ✅ 自动识别内销/外销文件类型
        ✅ 智能推荐对应格式的订单号列
        ✅ 订单号自动转为文本格式，防止科学计数法
        ✅ 支持两种格式混合匹配（汇总内销+明细外销/反之）
        """)
        st.markdown("---")
        st.markdown("### 版本信息")
        st.info("版本: v2.0.0（内销+外销通用）")
    
    # 主界面
    st.title("🔗 订单匹配工具（内销+外销通用）")
    st.markdown("支持内销/外销格式，根据订单号匹配汇总表和明细表数据")
    st.markdown("---")
    
    # 创建两列布局
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("📊 汇总表")
        summary_file = st.file_uploader(
            "上传汇总信息 Excel 文件（内销/外销均可）",
            type=['xlsx', 'xls'],
            key="summary_file"
        )
        if summary_file:
            st.success(f"✅ 已上传: {summary_file.name}")
    with col2:
        st.subheader("📋 明细表")
        detail_file = st.file_uploader(
            "上传明细信息 Excel 文件（内销/外销均可）",
            type=['xlsx', 'xls'],
            key="detail_file"
        )
        if detail_file:
            st.success(f"✅ 已上传: {detail_file.name}")
    
    # 读取数据和识别文件类型
    df_summary = None
    df_detail = None
    summary_col = None
    detail_col = None
    summary_columns = []
    detail_columns = []
    summary_file_type = None  # 'domestic'/'export'/None
    detail_file_type = None   # 'domestic'/'export'/None
    
    if summary_file:
        try:
            df_summary = pd.read_excel(summary_file)
            st.info(f"📊 汇总表: {len(df_summary)} 行, {len(df_summary.columns)} 列")
            # 识别汇总表文件类型
            summary_file_type = detect_file_type(df_summary, summary_file.name)
            if summary_file_type:
                st.success(f"📌 自动识别为: {'内销' if summary_file_type == 'domestic' else '外销'}格式")
            else:
                st.warning("⚠️ 无法自动识别文件类型，请手动确认订单号列")
        except Exception as e:
            st.error(f"❌ 读取汇总表失败: {e}")
    
    if detail_file:
        try:
            df_detail = pd.read_excel(detail_file)
            st.info(f"📋 明细表: {len(df_detail)} 行, {len(df_detail.columns)} 列")
            # 识别明细表文件类型
            detail_file_type = detect_file_type(df_detail, detail_file.name)
            if detail_file_type:
                st.success(f"📌 自动识别为: {'内销' if detail_file_type == 'domestic' else '外销'}格式")
            else:
                st.warning("⚠️ 无法自动识别文件类型，请手动确认订单号列")
        except Exception as e:
            st.error(f"❌ 读取明细表失败: {e}")
    
    st.markdown("---")
    
    # 选择订单号列（适配双格式）
    if df_summary is not None and df_detail is not None:
        st.subheader("🏷️ 选择订单号列")
        # 自动识别（传入文件类型，优化推荐精度）
        summary_recommended, summary_sorted = detect_column(df_summary, "汇总表", summary_file_type)
        detail_recommended, detail_sorted = detect_column(df_detail, "明细表", detail_file_type)
        
        col1, col2 = st.columns(2)
        with col1:
            st.markdown(f"**汇总表订单号列（{'内销' if summary_file_type == 'domestic' else '外销'}格式）：**")
            if summary_recommended:
                st.caption(f"💡 推荐列: **{summary_recommended}**")
            summary_col = st.selectbox(
                "选择汇总表中的订单号列",
                options=summary_sorted,
                index=0 if summary_sorted else None,
                key="summary_col_select"
            )
        with col2:
            st.markdown(f"**明细表订单号列（{'内销' if detail_file_type == 'domestic' else '外销'}格式）：**")
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
                    f"- 汇总表：**{summary_col}**（{'内销' if summary_file_type == 'domestic' else '外销'}格式）\n"
                    f"- 明细表：**{detail_col}**（{'内销' if detail_file_type == 'domestic' else '外销'}格式）")
        
        # 开始匹配按钮
        if st.button("🔗 开始匹配", type="primary", use_container_width=True):
            if not summary_col or not detail_col:
                st.error("❌ 请先选择订单号列！")
            else:
                with st.spinner("正在匹配数据..."):
                    try:
                        # 关键优化：订单号列转换为字符串（适配内销可能的特殊格式）
                        df_summary[summary_col] = df_summary[summary_col].astype(str).str.strip()
                        df_detail[detail_col] = df_detail[detail_col].astype(str).str.strip()
                        
                        # 匹配逻辑（兼容内销/外销订单号格式）
                        summary_orders = df_summary[summary_col].dropna().unique()
                        order_set = set(summary_orders)
                        matched = df_detail[df_detail[detail_col].isin(order_set)]
                        
                        if matched.empty:
                            st.warning("⚠️ 没有找到任何匹配的订单！")
                            # 显示不匹配原因分析
                            with st.expander("查看不匹配原因分析"):
                                st.markdown(f"- 汇总表订单号样本：{list(summary_orders[:5])}")
                                st.markdown(f"- 明细表订单号样本：{list(df_detail[detail_col].dropna().unique()[:5])}")
                                st.markdown("建议：确认订单号列选择正确，或检查订单号格式是否一致（如是否包含前缀/后缀）")
                        else:
                            # 显示匹配结果
                            st.success(f"✅ 匹配完成！找到 **{len(matched)}** 条匹配记录")
                            
                            # 显示统计信息
                            col1, col2, col3 = st.columns(3)
                            with col1:
                                st.metric("汇总表订单数", len(summary_orders))
                            with col2:
                                st.metric("明细表记录数", len(df_detail))
                            with col3:
                                st.metric("匹配记录数", len(matched))
                            
                            # 预览数据
                            st.subheader("📊 匹配结果预览")
                            st.dataframe(matched.head(20), use_container_width=True)
                            if len(matched) > 20:
                                st.caption(f"仅显示前 20 条记录，共 {len(matched)} 条")
                            
                            # 生成下载文件（保持原格式，订单号列设为文本）
                            output = io.BytesIO()
                            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                                matched.to_excel(writer, index=False, sheet_name='匹配结果')
                                workbook = writer.book
                                worksheet = writer.sheets['匹配结果']
                                
                                # 获取订单号列位置并设置为文本格式（防止科学计数法）
                                col_idx = matched.columns.get_loc(detail_col)
                                col_letter = col_num_to_letter(col_idx)
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