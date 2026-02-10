import streamlit as st
import pandas as pd
import random
import os

# ================= 1. Apple Pro 视觉引擎 (极致窄版 & 纯黑文字) =================
# 布局改为 centered
st.set_page_config(page_title="肆叁叁月季起名社", page_icon="💐", layout="centered")

st.markdown("""
<style>
    /* 引入思源宋体与 Inter */
    @import url('https://fonts.googleapis.com/css2?family=Noto+Serif+SC:wght@700;900&family=Inter:wght@400;600&display=swap');

    /* 1. 全局背景 */
    .stApp {
        background-color: #F5F5F7 !important;
        background-image: radial-gradient(circle at 50% 10%, #FFFFFF 0%, #E2E2E7 100%) !important;
        color: #000000 !important;
    }

    /* 2. 艺术标题区 */
    .header-box {
        padding-top: 80px;
        padding-bottom: 30px;
        text-align: center;
    }
    .artistic-title {
        font-family: 'Noto Serif SC', serif;
        font-size: 58px;
        font-weight: 900;
        color: #000000 !important;
        letter-spacing: 14px;
        margin-bottom: 10px;
    }
    .artistic-subtitle {
        font-family: 'Inter', sans-serif;
        color: #86868B !important;
        font-size: 13px;
        letter-spacing: 5px;
        text-transform: uppercase;
    }

    /* 3. 【核心修正】精致窄版控制台 (限制 650px) */
    div[data-testid="stVerticalBlockBorderWrapper"]:first-of-type {
        max-width: 650px !important;
        margin: 0 auto !important;
        background: rgba(255, 255, 255, 0.92) !important;
        backdrop-filter: blur(50px) saturate(180%) !important;
        -webkit-backdrop-filter: blur(50px) saturate(180%) !important;
        border: 2px solid #FFFFFF !important;
        border-radius: 36px !important;
        box-shadow: 0 20px 60px rgba(0,0,0,0.05), inset 0 1px 2px rgba(255,255,255,0.5) !important;
        padding: 45px !important;
    }

    /* 4. 强制列并行展示 (解决 3 列并排过宽问题) */
    div[data-testid="stHorizontalBlock"] {
        flex-direction: row !important;
        flex-wrap: nowrap !important;
        gap: 12px !important;
    }
    div[data-testid="column"] {
        min-width: 0 !important;
        flex: 1 1 0% !important;
    }

    /* 5. 强制纯黑文字 */
    p, span, label, div[data-testid="stMarkdownContainer"] p {
        color: #000000 !important;
        font-weight: 500 !important;
    }

    /* 6. 输入控件优化 */
    .stSelectbox div[data-baseweb="select"], .stTextInput input {
        background-color: #F0F0F2 !important;
        border: none !important;
        border-radius: 12px !important;
        color: #000000 !important;
    }

    /* 7. 苹果蓝按钮 */
    div.stButton { text-align: center; margin-top: 25px; }
    .stButton>button {
        background: #0071E3 !important;
        border-radius: 99px !important;
        color: white !important;
        padding: 10px 60px !important;
        font-size: 16px !important;
        border: none !important;
        transition: all 0.3s;
    }
    .stButton>button:hover { background: #0077ED !important; transform: scale(1.03); }

    /* 8. 结果号牌 (同步收窄) */
    .res-section {
        max-width: 650px;
        margin: 0 auto;
    }
    .res-card {
        background: #FFFFFF !important;
        border-radius: 20px !important;
        padding: 25px !important;
        text-align: center !important;
        border: 1px solid #E5E5E7 !important;
        margin-top: 15px !important;
        box-shadow: 0 10px 25px rgba(0,0,0,0.03) !important;
    }
    .res-text { font-size: 26px !important; font-weight: 700 !important; color: #000000 !important; letter-spacing: -0.5px !important; }

    /* 9. 字库艺术标题 */
    .lib-section {
        max-width: 650px !important;
        margin: 60px auto 140px auto !important;
        text-align: center !important;
    }
    .lib-title-art {
        font-family: 'Noto Serif SC', serif;
        font-size: 24px;
        font-weight: 900;
        color: #000000 !important;
        letter-spacing: 8px;
        margin-bottom: 20px;
    }
    .lib-tag {
        display: inline-block !important;
        background: white !important;
        padding: 3px 10px !important;
        border-radius: 6px !important;
        margin: 3px !important;
        font-size: 12px !important;
        color: #424245 !important;
        border: 1px solid #E5E5E7 !important;
    }

    /* 10. 强力置底页脚 */
    .fixed-footer {
        position: fixed !important;
        bottom: 0 !important;
        left: 0 !important;
        width: 100% !important;
        background: rgba(245, 245, 247, 0.95) !important;
        backdrop-filter: blur(15px) !important;
        text-align: center !important;
        padding: 18px 0 !important;
        border-top: 1px solid #D2D2D7 !important;
        color: #86868B !important;
        font-size: 13px !important;
        z-index: 9999 !important;
    }

    /* 隐藏杂项 */
    header, footer, [data-testid="stHeader"] { visibility: hidden !important; }
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

# ================= 3. UI 主体逻辑 =================

# 艺术大标题
st.markdown("""
<div class="header-box">
    <div class="artistic-title">命名工作站</div>
    <div class="artistic-subtitle">Pure Artistry for Every Rose</div>
</div>
""", unsafe_allow_html=True)

# 【核心参数面板】
with st.container(border=True):
    # 第一排 3 列
    c1, c2, c3 = st.columns([1, 1.2, 1.2])
    with c1: brand = st.text_input("品牌词", value="中农")
    with c2: color_cat = st.selectbox("核心色系", db["color"]["分类"].unique())
    with c3:
        scheme_opts = db["color"][db["color"]["分类"] == color_cat]["方案"].unique().tolist()
        scheme_sel = st.selectbox("方案风格", scheme_opts)
    
    # 第二排 2 列
    c4, c5 = st.columns(2)
    with c4: pre_cat = st.selectbox("前缀性状 (可选)", ["(无)"] + db["prefix"]["性状名称"].unique().tolist())
    with c5: suf_cat = st.selectbox("后缀性状 (必选)", db["suffix"]["性状名称"].unique())
    
    # 第三排并排布局
    c6, c7, c8 = st.columns([1, 1.2, 1.2])
    with c6: attr_mode = st.radio("属性偏好", ["表型", "意象"], horizontal=True)
    with c7: tail_cat = st.selectbox("尾缀色系 (可选)", ["(无)"] + db["color"]["分类"].unique().tolist())
    with c8:
        if tail_cat != "(无)":
            t_schemes = db["color"][db["color"]["分类"] == tail_cat]["方案"].unique().tolist()
            tail_scheme = st.selectbox("尾缀风格", t_schemes)
        else: tail_scheme = None
    
    run_gen = st.button("智能生成方案")

# ================= 4. 生成结果 =================

def generate_logic(count=10):
    core_chars = db["color"][(db["color"]["分类"] == color_cat) & (db["color"]["方案"] == scheme_sel)]["汉字"].tolist()
    p_pool = db["prefix"][db["prefix"]["性状名称"] == pre_cat]["汉字"].tolist() if pre_cat != "(无)" else []
    s_pool = db["suffix"][(db["suffix"]["性状名称"] == suf_cat) & (db["suffix"]["属性"] == attr_mode)]["汉字"].tolist()
    if not s_pool: s_pool = db["suffix"][db["suffix"]["性状名称"] == suf_cat]["汉字"].tolist()
    t_pool = db["color"][(db["color"]["分类"] == tail_cat) & (db["color"]["方案"] == tail_scheme)]["汉字"].tolist() if tail_cat != "(无)" and tail_scheme else []

    results = []
    for _ in range(count):
        c = random.choice(core_chars) if core_chars else ""
        p = random.choice(p_pool) if p_pool else ""
        s = random.choice(s_pool) if s_pool else ""
        t = random.choice(t_pool) if t_pool else ""
        results.append(f"'{brand} {p}{c}{s}{t}'")
    return results

if run_gen:
    names = generate_logic(10)
    st.markdown('<div class="res-section">', unsafe_allow_html=True)
    res_cols = st.columns(2)
    for idx, full_name in enumerate(names):
        with res_cols[idx % 2]:
            st.markdown(f"""<div class="res-card"><div style="font-size:10px; color:#0071E3; font-weight:700;">NO.{idx+1:02d}</div><div class="res-text">{full_name}</div></div>""", unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)
else:
    st.markdown('<p style="text-align:center; color:#86868B; margin-top:30px; font-size:14px;">点击按钮开启 AI 灵感推荐</p>', unsafe_allow_html=True)

# ================= 5. 字库展示 (艺术标题版) =================
st.markdown('<div class="lib-section">', unsafe_allow_html=True)
st.markdown('<div class="lib-title-art">字库全览</div>', unsafe_allow_html=True)
t1, t2, t3 = st.tabs(["🌈 核心色", "🏷️ 修饰前缀", "✨ 性状后缀"])
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

# ================= 6. 固定脚注 (按要求更新) =================
st.markdown("""
<div class="fixed-footer">
    © 2026 肆叁叁月季起名社
</div>
""", unsafe_allow_html=True)