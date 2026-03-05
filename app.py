import streamlit as st
from pages.batch_update import render_batch_update
from pages.field_extraction import render_field_extraction

# --- 1. 页面配置与 CSS ---
st.set_page_config(page_title="小雷Excel批量助手", page_icon="⚡", layout="wide")

st.markdown("""
    <style>
    /* 全局背景与字体 */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&display=swap');
    
    .main {
        background-color: #fcfdfe;
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
    }
    
    .block-container { 
        padding-top: 2rem !important; 
        max-width: 1000px !important;
    }
    
    header {visibility: hidden;}
    footer {visibility: hidden;}

    /* 标题样式 */
    .main-title {
        font-size: 1.8rem;
        font-weight: 700;
        color: #1e293b;
        margin-bottom: 0.5rem;
        text-align: center;
        background: linear-gradient(90deg, #3b82f6, #06b6d4);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
    }
    .sub-title {
        font-size: 0.9rem;
        color: #64748b;
        text-align: center;
        margin-bottom: 2rem;
    }

    /* 卡片通用样式 */
    .stFileUploader section {
        border: 2px dashed #e2e8f0 !important;
        background-color: #ffffff !important;
        padding: 1.5rem !important;
        border-radius: 12px !important;
        transition: all 0.3s ease;
    }
    .stFileUploader section:hover {
        border-color: #3b82f6 !important;
        background-color: #f8fafc !important;
    }

    .upload-card {
        background: #ffffff;
        border: 1px solid #f1f5f9;
        padding: 20px;
        border-radius: 16px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05), 0 2px 4px -1px rgba(0, 0, 0, 0.03);
        margin-bottom: 20px;
    }

    .config-card {
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        padding: 18px;
        border-radius: 12px;
        margin-bottom: 15px;
    }

    .personal-box {
        padding: 15px;
        background: #ffffff;
        border-radius: 12px;
        margin-top: 10px;
        border: 1px solid #f1f5f9;
        box-shadow: inset 0 2px 4px 0 rgba(0, 0, 0, 0.02);
        border-left: 4px solid #3b82f6;
    }

    .download-bar {
        background: #ecfdf5;
        border: 1px solid #10b981;
        padding: 15px 20px;
        border-radius: 12px;
        margin: 20px 0;
        display: flex;
        align-items: center;
        justify-content: space-between;
        color: #065f46;
        font-weight: 500;
    }

    /* 按钮美化 */
    .stButton > button {
        border-radius: 10px !important;
        padding: 0.5rem 1rem !important;
        font-weight: 600 !important;
        transition: all 0.2s ease !important;
    }
    .stButton > button:hover {
        transform: translateY(-1px);
        box-shadow: 0 4px 12px rgba(59, 130, 246, 0.25);
    }
    
    /* 侧边栏/输入框样式 */
    h5 {
        margin-bottom: 0.8rem !important;
        font-weight: 600 !important;
        font-size: 0.95rem !important;
        color: #1e293b !important;
        display: flex;
        align-items: center;
        gap: 8px;
    }
    
    .stExpander {
        border-radius: 12px !important;
        border: 1px solid #f1f5f9 !important;
        background: #ffffff !important;
    }
    </style>
    """, unsafe_allow_html=True)

if 'batch_results' not in st.session_state:
    st.session_state.batch_results = []
if 'extract_results' not in st.session_state:
    st.session_state.extract_results = []

# --- 3. 界面交互 ---
st.markdown('<div class="main-title">⚡ 小雷 Excel 批量助手</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">高效、简单、本地处理的 Excel 办公利器</div>', unsafe_allow_html=True)

tab1, tab2 = st.tabs(["🔄 批量关联更新", "✨ 批量字段提取"])

with tab1:
    render_batch_update()

with tab2:
    render_field_extraction()

# --- 4. 底部声明 ---
st.markdown("---")
st.markdown(
    '<div style="text-align: center; color: #94a3b8; font-size: 0.8rem; padding: 20px;">'
    '小雷 Excel 助手 · 本地处理更安全 · 2024'
    '</div>', 
    unsafe_allow_html=True
)
