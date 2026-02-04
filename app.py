import streamlit as st
import pandas as pd
import random
import os

# ================= 页面配置 =================
st.set_page_config(
    page_title="RoseNamer Elite",
    page_icon="🌹",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# ================= 极简苹果风 UI =================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Serif+SC:wght@700;900&family=Inter:wght@400;600&display=swap');

html, body, [data-testid="stAppViewContainer"] {
    background-color: #F5F5F7 !important;
    color: #1D1D1F !important;
}

.header-box {padding-top:80px;text-align:center;}
.artistic-title {
    font-family:'Noto Serif SC',serif;
    font-size:54px;font-weight:900;letter-spacing:16px;
}
.artistic-subtitle {
    font-family:'Inter',sans-serif;
    color:#86868B;font-size:13px;letter-spacing:6px;
}

.panel {
    max-width:580px;margin:0 auto;
    background:#FFFFFF;border-radius:28px;
    border:1px solid #D2D2D7;padding:40px;
    box-shadow:0 20px 40px rgba(0,0,0,0.05);
}

.stTextInput input, .stSelectbox div[data-baseweb="select"] {
    background:#F5F5F7 !important;
    border-radius:12px;border:1px solid #E5E5E7;
}

div[data-baseweb="popover"], div[role="listbox"] {
    background:#FFFFFF !important;color:#1D1D1F !important;
}

.stButton>button {
    background:#0071E3;color:white;border-radius:99px;
    padding:10px 50px;font-size:16px;border:none;
}
.stButton>button:hover {background:#0077ED;transform:scale(1.02);}

.res-card {
    background:#FFF;border-radius:18px;padding:20px;
    border:1px solid #E5E5E7;margin-top:12px;
    text-align:center;box-shadow:0 10px 20px rgba(0,0,0,0.03);
}

.fixed-footer {
    position:fixed;bottom:0;width:100%;
    text-align:center;padding:15px 0;
    border-top:1px solid #D2D2D7;
    color:#86868B;font-size:12px;
    background:rgba(245,245,247,0.95);
}
</style>
""", unsafe_allow_html=True)

# ================= 数据加载 =================
EXCEL_FILE = "rose_data.xlsx"

@st.cache_data
def load_db(file):
    if not os.path.exists(file):
        return None
    xls = pd.ExcelFile(file)
    return {
        "color": pd.read_excel(xls, "色核库"),
        "suffix": pd.read_excel(xls, "后缀映射"),
        "prefix": pd.read_excel(xls, "前缀库")
    }

db = load_db(EXCEL_FILE)

if db is None:
    st.error("❌ 找不到 rose_data.xlsx")
    st.stop()

# ================= 预处理加速 =================
@st.cache_data
def preprocess_db(db):
    return {
        "color_map": {k: v["汉字"].tolist() for k, v in db["color"].groupby("分类")},
        "prefix_map": {k: v["汉字"].tolist() for k, v in db["prefix"].groupby("性状名称")},
        "suffix_map": {(k,a): v["汉字"].tolist() for (k,a), v in db["suffix"].groupby(["性状名称","属性"])}
    }

maps = preprocess_db(db)

# ================= 标题 =================
st.markdown("""
<div class="header-box">
<div class="artistic-title">命名工作站</div>
<div class="artistic-subtitle">PURE ARTISTRY FOR EVERY ROSE</div>
</div>
""", unsafe_allow_html=True)

# ================= 控制台 =================
st.markdown('<div class="panel">', unsafe_allow_html=True)

c1, c2 = st.columns(2)
with c1:
    brand = st.text_input("品牌词", "中农")
with c2:
    color_cat = st.selectbox("核心色系", list(maps["color_map"].keys()))

c3, c4 = st.columns(2)
with c3:
    pre_cat = st.selectbox("前缀性状", ["(无)"] + list(maps["prefix_map"].keys()))
with c4:
    suf_cat = st.selectbox("后缀性状", list({k for k,a in maps["suffix_map"].keys()}))

c5, c6 = st.columns(2)
with c5:
    attr_mode = st.radio("属性偏好", ["表型", "意象"], horizontal=True)
with c6:
    tail_cat = st.selectbox("第二色核", ["(无)"] + list(maps["color_map"].keys()))

run_gen = st.button("智能生成方案")
st.markdown('</div>', unsafe_allow_html=True)

# ================= 安全生成 =================
def safe_choice(lst, default=""):
    return random.choice(lst) if lst else default

def generate_logic(n=10):
    core = maps["color_map"].get(color_cat, [])
    pre = maps["prefix_map"].get(pre_cat, []) if pre_cat != "(无)" else []
    suf = maps["suffix_map"].get((suf_cat, attr_mode), []) or \
          [v for (k,a),v in maps["suffix_map"].items() if k==suf_cat][0]
    tail = maps["color_map"].get(tail_cat, []) if tail_cat != "(无)" else []

    results = []
    for _ in range(n):
        name = f"{brand} {safe_choice(pre)}{safe_choice(core,'玫')}{safe_choice(suf,'韵')}{safe_choice(tail)}"
        results.append(name)
    return results

# ================= 结果展示 =================
if run_gen:
    names = generate_logic()
    cols = st.columns(2)
    for i, name in enumerate(names):
        with cols[i % 2]:
            st.markdown(f"<div class='res-card'><b>NO.{i+1:02d}</b><br>{name}</div>", unsafe_allow_html=True)

# ================= 页脚 =================
st.markdown("<div class='fixed-footer'>© 2026 肆叁叁月季起名社</div>", unsafe_allow_html=True)
