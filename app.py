"""
A23 智能填表系统 - 主程序
基于大语言模型的文档理解与多源数据融合系统
"""

import streamlit as st
import pandas as pd
from docx import Document
import openai
import os
from tempfile import NamedTemporaryFile
import asyncio
import aiohttp
from typing import List, Dict
import time

# 页面配置（必须放在第一行）
st.set_page_config(
    page_title="A23 智能填表系统",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==================== 标题区域 ====================
st.title("📄 A23 智能填表系统")
st.markdown("*基于大语言模型的文档理解与多源数据融合系统*")
st.markdown("---")

# ==================== 侧边栏配置 ====================
with st.sidebar:
    st.header("⚙️ 系统配置")
    
    # API配置
    st.subheader("🤖 大模型配置")
    api_key = st.text_input("API Key", type="password", 
                            help="从阿里云/OpenAI获取的API密钥")
    
    api_base = st.text_input("API Base URL", 
                            value="https://dashscope.aliyuncs.com/compatible-mode/v1",
                            help="默认是阿里云通义千问的地址")
    
    model = st.selectbox("选择模型", 
                        ["qwen-plus", "qwen-turbo", "gpt-3.5-turbo", "自定义"],
                        help="qwen是阿里云模型，gpt需要OpenAI密钥")
    
    if model == "自定义":
        model = st.text_input("输入模型名称", value="qwen-plus")
    
    # 高级配置
    with st.expander("⚡ 高级配置"):
        max_concurrent = st.slider("最大并发数", 1, 20, 5,
                                  help="并发请求数越大越快，但可能被API限制")
        temperature = st.slider("Temperature", 0.0, 1.0, 0.1,
                               help="值越大回答越随机，值越小越确定")
    
    st.markdown("---")
    st.info("📌 上传文档和Excel模板，系统会自动填表")
    st.caption("第十七届服创大赛 · A23赛题")

# ==================== 主界面 ====================
col1, col2 = st.columns(2)

# 左侧：文档上传
with col1:
    st.subheader("📁 1. 上传文档")
    st.caption("支持格式：txt / docx / md / xlsx（作为数据源）")
    
    docs = st.file_uploader(
        "选择文档文件（可多选）",
        accept_multiple_files=True,
        type=['txt', 'docx', 'md', 'xlsx']
    )
    
    if docs:
        st.success(f"✅ 已上传 {len(docs)} 个文件")
        for doc in docs:
            file_size = len(doc.getvalue()) / 1024  # KB
            st.text(f"  • {doc.name} ({file_size:.1f} KB)")

# 右侧：表格上传
with col2:
    st.subheader("📊 2. 上传Excel模板")
    st.caption("模板格式：第一行为表头，第一列为待填项标识")
    
    template = st.file_uploader(
        "选择Excel模板文件",
        type=['xlsx']
    )
    
    if template:
        st.success(f"✅ 已上传模板：{template.name}")
        
        # 预览模板内容
        try:
            # 临时保存文件
            with NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
                tmp.write(template.getvalue())
                tmp_path = tmp.name
            
            # 读取Excel
            df = pd.read_excel(tmp_path)
            
            # 清理临时文件
            os.unlink(tmp_path)
            
            st.write("📋 模板预览：")
            st.dataframe(df.head(5), use_container_width=True)
            
            # 显示表格结构
            st.caption(f"表格大小：{df.shape[0]}行 × {df.shape[1]}列")
            
            # 保存到session state供后续使用
            st.session_state['template_df'] = df
            st.session_state['template_name'] = template.name
            
        except Exception as e:
            st.error(f"❌ 无法读取Excel文件：{str(e)}")

# ==================== 处理按钮 ====================
st.markdown("---")
col_btn, col_status = st.columns([1, 3])

with col_btn:
    start_btn = st.button("🚀 开始智能填表", type="primary", use_container_width=True)

with col_status:
    status_placeholder = st.empty()

# ==================== 处理逻辑 ====================
if start_btn:
    # 检查必要输入
    if not docs:
        st.error("❌ 请先上传文档")
        st.stop()
    if not template:
        st.error("❌ 请先上传Excel模板")
        st.stop()
    if not api_key:
        st.error("❌ 请在侧边栏配置API Key")
        st.stop()
    
    # 开始计时
    start_time = time.time()
    
    # 显示进度
    progress_bar = st.progress(0, text="准备开始...")
    status_text = st.empty()
    
    try:
        # ========== 第1步：解析文档 ==========
        status_text.text("📖 正在解析文档...")
        progress_bar.progress(10, text="解析文档中...")
        
        all_text = ""
        for i, doc in enumerate(docs):
            # 这里先简单处理，后面会完善
            status_text.text(f"  处理 {doc.name}...")
            # 模拟解析时间
            time.sleep(0.5)
            all_text += f"[来自文件：{doc.name}]\n"
        
        # ========== 第2步：解析模板 ==========
        status_text.text("📋 正在分析表格结构...")
        progress_bar.progress(30, text="分析表格中...")
        
        df = st.session_state.get('template_df')
        if df is None:
            st.error("表格数据丢失，请重新上传")
            st.stop()
        
        # 获取表头
        headers = df.columns.tolist()
        status_text.text(f"  发现 {len(headers)} 列表头：{', '.join(headers[:3])}...")
        
        # ========== 第3步：逐行填表（模拟） ==========
        status_text.text("🤖 AI正在填表...")
        progress_bar.progress(50, text="AI处理中...")
        
        # 创建结果DataFrame
        result_df = df.copy()
        
        # 逐行处理
        total_rows = len(df)
        for idx in range(total_rows):
            progress = 50 + int(40 * (idx + 1) / total_rows)
            progress_bar.progress(progress, text=f"正在处理第 {idx+1}/{total_rows} 行...")
            
            # 这里先填模拟数据，后面换成真正的AI调用
            result_df.iloc[idx, 1:] = f"AI生成内容_{idx+1}"
            
            # 稍微停顿，看起来在运行
            time.sleep(0.3)
        
        # ========== 第4步：生成结果 ==========
        status_text.text("💾 正在生成结果文件...")
        progress_bar.progress(95, text="生成文件中...")
        
        # 保存结果到临时文件
        with NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
            result_df.to_excel(tmp.name, index=False)
            result_path = tmp.name
        
        # 计算用时
        elapsed_time = time.time() - start_time
        
        # ========== 显示结果 ==========
        progress_bar.progress(100, text="完成！")
        status_text.text("")
        
        st.success(f"✅ 填表完成！用时 {elapsed_time:.1f} 秒")
        
        # 显示结果预览
        st.subheader("📊 填表结果预览")
        st.dataframe(result_df.head(10), use_container_width=True)
        
        # 下载按钮
        with open(result_path, 'rb') as f:
            st.download_button(
                label="📥 下载填好的Excel",
                data=f,
                file_name=f"filled_{st.session_state.get('template_name', 'result.xlsx')}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        
        # 清理临时文件
        os.unlink(result_path)
        
    except Exception as e:
        st.error(f"❌ 处理过程中出错：{str(e)}")
        progress_bar.empty()

# ==================== 底部 ====================
st.markdown("---")
st.caption("© 2025 A23项目组 · 第十七届服创大赛")