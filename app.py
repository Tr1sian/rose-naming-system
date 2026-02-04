import streamlit as st
import pandas as pd
import random
import os

# ================= 1. 苹果美学 Pro (强制浅色主题 & 布局锁定) =================
st.set_page_config(page_title="RoseNamer Elite", page_icon="🍎", layout="centered")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Serif+SC:wght@700;900&family=Inter:wght@300;400;600&display=swap');

    /* 1. 强制重置 Streamlit 核心变量：彻底杀掉深色模式 */
    :root {
        --primary-color: #0071E3;
        --background-color: #F5F5F7;
        --secondary-background-color: #FFFFFF;
        --text-color: #1D1D1F;
    }

    /* 2. 全局背景强制还原 */
    .stApp {
        background: #F5F5F7 !important;
        background-image: radial-gradient(circle at 50% 10%, #FFFFFF 0%, #E2E2E7 100%) !important;
    }

    /* 3. 标题艺术化排版 */
    .header-box {
        padding-top: 80px;
        padding-bottom: 20px;
        text-align: center;
    }
    .artistic-title {
        font-family: 'Noto Serif SC', serif;
        font-size: 54px;
        font-weight: 900;
        color: #1D1D1F !important;
        letter-spacing: 16px;
        margin-bottom: 5px;
    }
    .artistic-subtitle {
        font-family: 'Inter', sans-serif;
        color: #86868B !important;
        font-size: 13px;
        letter-spacing: 6px;
        text-transform: uppercase;
    }

    /* 4. 【核心修复】纯白立体玻璃参数面板 (解决黑框问题) */
    div[data-testid="stVerticalBlockBorderWrapper"]:first-of-type {
        max-width: 580px !important;
        margin: 0 auto !important;
        /* 强制纯白，高不透明度以盖住深色干扰 */
        background: rgba(255, 255, 255, 0.98) !important;
        backdrop-filter: blur(40px) !important;
        -webkit-backdrop-filter: blur(40px) !important;
        border: 2px solid #FFFFFF !important;
        border-radius: 32px !important;
        box-shadow: 0 20px 60px rgba(0,0,0,0.05), inset 0 1px 2px rgba(255,255,255,0.8) !important;
        padding: 40px !important;
    }

    /* 5. 强制所有内部文字颜色为黑色 */
    p, span, label, div {
        color: #1D1D1F !important;
    }
    
    /* 强制副标题和提示语颜色为灰色 */
    .artistic-subtitle, .run-hint, .fixed-footer {
        color: #86868B !important;
    }

    /* 6. 强力修正输入控件：解决变黑、无法对齐的问题 */
    /* 下拉框与输入框强制白底黑字 */
    div[data-baseweb="select"], div[data-baseweb="input"], input {
        background-color: #F5F5F7 !important;
        color: #1D1D1F !important;
        border: none !important;
        border-radius: 12px !important;
    }

    /* 修正下拉框内部箭头和文本 */
    div[data-baseweb="select"] * {
        color: #1D1D1F !important;
        background: transparent !important;
    }

    /* 7. 锁定并排布局：无论屏幕多窄都必须一行 */
    div[data-testid="stHorizontalBlock"] {
        flex-direction: row !important;
        flex-wrap: nowrap !important;
        gap: 1rem !important;
    }
    div[data-testid="column"] {
        width: 50% !important;
        min-width: 0 !important;
        flex: 1 1 auto !important;
    }

    /* 压缩多余的垂直间距 */
    div[data-testid="stVerticalBlock"] > div {
        padding-top: 0rem !important;
        padding-bottom: 0.5rem !important;
    }

    /* 8. 按钮苹果蓝美化 */
    div.stButton { text-align: center; margin-top: 15px; }
    .stButton>button {
        background: #0071E3 !important;
        border: none !important;
        border-radius: 99px !important;
        color: white !important;
        padding: 8px 50px !important;
        font-size: 16px !important;
        transition: 0.3s;
    }
    .stButton>button:hover { background: #0077ED !important; transform: scale(1.03); }

    /* 9. 结果展示卡片 */
    .res-card {
        background: #FFFFFF !important;
        border-radius: 18px !important;
        padding: 20px !important;
        text-align: center !important;
        border: 1px solid #E5E5E7 !important;
        margin-top: 12px !important;
        box-shadow: 0 10px 25px rgba(0,0,0,0.03) !important;
    }
    .res-text { font-size: 24px !important; font-weight: 700 !important; color: #1D1D1F !important; }

    /* 10. 字库展示区轻盈化 */
    .library-section {
        max-width: 600px !important;
        margin: 40px auto 120px auto !important;
        text-align: center !important;
    }
    .lib-tag {
        display: inline-block !important;
        background: white !important;
        padding: 3px 8px !important;
        border-radius: 6px !important;
        margin: 2px !important;
        color: #6E6E73 !important;
        border: 1px solid #E5E5E7 !important;
        font-size: 12px !important;
    }

    /* 固定页脚 */
    .fixed-footer {
        position: fixed; bottom: 0; left: 0; width: 100%;
        background: rgba(245, 245, 247, 0.95);
        backdrop-filter: blur(15px);
        text-align: center; padding: 15px 0;
        border-top: 1px solid #D2D2D7; font-size: 12px; z-index: 1000;
    }

    header, footer, [data-testid="stHeader"] { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

# ================= 2. 数据加载 =================
EXCEL_FILE = "rose_data.xlsx"

@st.cache_data
def load_db(file):
    if not os.path.exists(file): return None
    xls = pd.ExcelFile(file)
    return {
        "color": pd.read_excel(xls, "色核库"),
        "suffix": pd.read_excel(xls, "后缀映射"),
        "prefix": pd.read_excel(xls, "前缀库")
    }

db = load_db(EXCEL_FILE)

# ================= 3. UI 交互 =================

# 标题
st.markdown("""
<div class="header-box">
    <div class="artistic-title">命名工作站</div>
    <div class="artistic-subtitle">Pure Artistry for Every Rose</div>
</div>
""", unsafe_allow_html=True)

# 参数控制面板
with st.container(border=True):
    c1, c2 = st.columns(2)
    with c1: brand = st.text_input("品牌词", value="中农")
    with c2: color_cat = st.selectbox("核心色系", db["color"]["分类"].unique())
    
    c3, c4 = st.columns(2)
    with c3: pre_cat = st.selectbox("前缀性状 (可选)", ["(无)"] + db["prefix"]["性状名称"].unique().tolist())
    with c4: suf_cat = st.selectbox("后缀性状 (必选)", db["suffix"]["性状名称"].unique())
    
    c5, c6 = st.columns(2)
    with c5: attr_mode = st.radio("属性偏好", ["表型", "意象"], horizontal=True)
    with c6: tail_cat = st.selectbox("第二色核 (可选)", ["(无)"] + db["color"]["分类"].unique().tolist())
    
    run_gen = st.button("智能生成方案")

# ================= 4. 条件渲染 (结果或提示) =================

def generate_logic(count=10):
    core_chars = db["color"][db["color"]["分类"] == color_cat]["汉字"].tolist()
    p_pool = db["prefix"][db["prefix"]["性状名称"] == pre_cat]["汉字"].tolist() if pre_cat != "(无)" else []
    s_pool = db["suffix"][(db["suffix"]["性状名称"] == suf_cat) & (db["suffix"]["属性"] == attr_mode)]["汉字"].tolist()
    if not s_pool: s_pool = db["suffix"][db["suffix"]["性状名称"] == suf_cat]["汉字"].tolist()
    t_pool = db["color"][db["color"]["分类"] == tail_cat]["汉字"].tolist() if tail_cat != "(无)" else []
    results = []
    for _ in range(count):
        c = random.choice(core_chars); p = random.choice(p_pool) if p_pool else ""; s = random.choice(s_pool); t = random.choice(t_pool) if t_pool else ""
        results.append(f"'{brand} {p}{c}{s}{t}'")
    return results

if run_gen:
    names = generate_logic(10)
    st.markdown('<div style="max-width:580px; margin: 0 auto; padding-top:10px;">', unsafe_allow_html=True)
    res_cols = st.columns(2)
    for idx, full_name in enumerate(names):
        with res_cols[idx % 2]:
            st.markdown(f"""<div class="res-card"><div style="font-size:10px; color:#0071E3; font-weight:700;">NO.{idx+1:02d}</div><div class="res-text">{full_name}</div></div>""", unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)
else:
    st.markdown('<p class="run-hint">点击上方 <b>Run</b> 按钮以获得 10 组预选名称建议</p>', unsafe_allow_html=True)

# ================= 5. 字库全览 =================
st.markdown('<div class="library-section">', unsafe_allow_html=True)
st.markdown('<div style="font-family:serif; font-size:20px; font-weight:700; margin-bottom:15px; color:#1D1D1F;">字库全览</div>', unsafe_allow_html=True)
t1, t2, t3 = st.tabs(["🌈核心色", "修饰前缀", "性状后缀"])
with t1:
    for cat in db["color"]["分类"].unique():
        chars = db["color"][db["color"]["分类"] == cat]["汉字"].unique()
        st.markdown(f"**{cat}**：{' '.join([f'<span class=\"lib-tag\">{c}</span>' for c in chars])}", unsafe_allow_html=True)
with t2:
    for cat in db["prefix"]["性状名称"].unique():
        chars = db["prefix"][db["prefix"]["性状名称"] == cat]["汉字"].unique()
        st.markdown(f"**{cat}**：{' '.join([f'<span class=\"lib-tag\">{c}</span>' for c in chars])}", unsafe_allow_html=True)
with t3:
    for cat in db["suffix"]["性状名称"].unique():
        chars = db["suffix"][db["suffix"]["性状名称"] == cat]["汉字"].tolist()
        st.markdown(f"**{cat}**：{' '.join([f'<span class=\"lib-tag\">{c}</span>' for c in chars])}", unsafe_allow_html=True)
st.markdown('</div>', unsafe_allow_html=True)

# 固定页脚
st.markdown("""<div class="fixed-footer">© 2026 肆叁叁月季起名社</div>""", unsafe_allow_html=True)