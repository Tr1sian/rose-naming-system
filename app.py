import streamlit as st
import pandas as pd
import random
import os

# ================= 1. Apple Pro 双模玻璃美学 (CSS) =================
st.set_page_config(page_title="RoseNamer Pro", page_icon="🍎", layout="wide")

st.markdown("""
<style>
    /* 引入思源宋体和 Inter */
    @import url('https://fonts.googleapis.com/css2?family=Noto+Serif+SC:wght@700;900&family=Inter:wght@300;400;600&display=swap');

    /* 苹果风格：响应式毛玻璃面板 (自动适配深色/浅色模式) */
    [data-testid="stVerticalBlockBorderWrapper"] {
        /* 背景色使用半透明，让底层背景透过来 */
        background: rgba(255, 255, 255, 0.1) !important;
        backdrop-filter: blur(40px) saturate(180%) !important;
        -webkit-backdrop-filter: blur(40px) saturate(180%) !important;
        
        /* 动态边框：根据背景自动适应 */
        border: 1px solid rgba(255, 255, 255, 0.2) !important;
        border-radius: 32px !important;
        
        /* 深度阴影 */
        box-shadow: 0 20px 50px rgba(0,0,0,0.1) !important;
        
        max-width: 680px !important;
        margin: 0 auto !important;
        padding: 50px !important;
    }

    /* 标题艺术化 */
    .header-box {
        padding-top: 60px;
        padding-bottom: 40px;
        text-align: center;
    }
    .artistic-title {
        font-family: 'Noto Serif SC', serif;
        font-size: 68px;
        font-weight: 900;
        letter-spacing: 16px;
        margin-bottom: 10px;
        /* 渐变字色 */
        background: linear-gradient(180deg, #5e5e5e 0%, #a1a1a1 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
    }

    /* 强制列并行展示 (解决堆叠问题) */
    div[data-testid="stHorizontalBlock"] {
        flex-direction: row !important;
        flex-wrap: nowrap !important;
        gap: 15px !important;
    }

    /* 按钮样式：苹果经典蓝色 */
    .stButton>button {
        background: #0071E3 !important;
        color: white !important;
        border-radius: 99px !important;
        padding: 10px 60px !important;
        font-size: 17px !important;
        border: none !important;
        width: 100% !important;
        transition: 0.3s cubic-bezier(0.4, 0, 0.2, 1);
    }
    .stButton>button:hover {
        background: #0077ED !important;
        transform: scale(1.02);
        box-shadow: 0 8px 20px rgba(0,113,227,0.4);
    }

    /* 结果号牌 */
    .res-card {
        background: rgba(255, 255, 255, 0.05);
        backdrop-filter: blur(10px);
        border-radius: 20px;
        padding: 30px;
        text-align: center;
        border: 1px solid rgba(255, 255, 255, 0.1);
        margin-top: 20px;
    }
    .res-text { font-size: 32px; font-weight: 700; letter-spacing: -1px; }

    /* 字库全览轻量化 */
    .lib-tag {
        display: inline-block;
        padding: 4px 12px;
        border-radius: 8px;
        margin: 4px;
        font-size: 13px;
        background: rgba(255, 255, 255, 0.1);
        border: 1px solid rgba(255, 255, 255, 0.1);
    }

    /* 隐藏原生干扰 */
    header, footer, [data-testid="stHeader"] { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

# ================= 2. 数据加载逻辑 =================
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

# ================= 3. UI 交互界面 =================

st.markdown("""
<div class="header-box">
    <div class="artistic-title">命名工作站</div>
    <div class="artistic-subtitle">CRAFTING SOULS FOR EVERY ROSE</div>
</div>
""", unsafe_allow_html=True)

with st.container(border=True):
    # 第一排：品牌与色核分级
    c1, c2, c3 = st.columns([1, 1.2, 1.2])
    with c1:
        brand = st.text_input("品牌词", value="中农")
    with c2:
        color_cat = st.selectbox("核心色系", db["color"]["分类"].unique())
    with c3:
        # 新增：方案类型二级联动
        scheme_opts = db["color"][db["color"]["分类"] == color_cat]["方案"].unique().tolist()
        scheme_sel = st.selectbox("方案风格", scheme_opts)

    # 第二排：前缀与后缀分类
    c4, c5 = st.columns(2)
    with c4:
        pre_cat = st.selectbox("前缀性状 (可选)", ["(无)"] + db["prefix"]["性状名称"].unique().tolist())
    with c5:
        suf_cat = st.selectbox("后缀性状 (必选)", db["suffix"]["性状名称"].unique())

    # 第三排：后缀偏好与尾缀联动
    c6, c7, c8 = st.columns([1, 1.2, 1.2])
    with c6:
        attr_mode = st.radio("后缀意境", ["表型", "意象"], horizontal=True)
    with c7:
        tail_cat = st.selectbox("第二色核 (可选)", ["(无)"] + db["color"]["分类"].unique().tolist())
    with c8:
        if tail_cat != "(无)":
            # 尾缀方案联动
            t_schemes = db["color"][db["color"]["分类"] == tail_cat]["方案"].unique().tolist()
            tail_scheme = st.selectbox("尾缀风格", t_schemes)
        else:
            tail_scheme = None

    run_gen = st.button("智能生成备选方案")

# ================= 4. 结果生成逻辑 =================

def generate_logic(count=10):
    # 核心色池：根据色系和方案双重过滤
    core_chars = db["color"][(db["color"]["分类"] == color_cat) & (db["color"]["方案"] == scheme_sel)]["汉字"].tolist()
    
    # 前缀池
    p_pool = db["prefix"][db["prefix"]["性状名称"] == pre_cat]["汉字"].tolist() if pre_cat != "(无)" else []
    
    # 后缀池
    s_pool = db["suffix"][(db["suffix"]["性状名称"] == suf_cat) & (db["suffix"]["属性"] == attr_mode)]["汉字"].tolist()
    if not s_pool: s_pool = db["suffix"][db["suffix"]["性状名称"] == suf_cat]["汉字"].tolist()
    
    # 尾缀池
    t_pool = []
    if tail_cat != "(无)" and tail_scheme:
        t_pool = db["color"][(db["color"]["分类"] == tail_cat) & (db["color"]["方案"] == tail_scheme)]["汉字"].tolist()

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
    st.markdown('<div style="max-width:550px; margin: 0 auto; padding-top:20px;">', unsafe_allow_html=True)
    res_cols = st.columns(2)
    for idx, full_name in enumerate(names):
        with res_cols[idx % 2]:
            st.markdown(f"""<div class="res-card"><div style="font-size:10px; color:#0071E3; font-weight:700;">SELECTION {idx+1:02d}</div><div class="res-text">{full_name}</div></div>""", unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)
else:
    st.markdown('<p style="text-align:center; color:#86868B; margin-top:30px; letter-spacing:2px;">点击上方按钮开启 AI 灵感推荐</p>', unsafe_allow_html=True)

# ================= 5. 字库展示 (置底) =================
st.write("")
with st.expander("📚 点击查看当前分类下的全部字库"):
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

# 6. 固定页脚
st.markdown("""<div class="fixed-footer">月季起名社 &nbsp; | &nbsp; © 2024 中农育种工作站 &nbsp; | &nbsp; 遵循《国际栽培植物命名法规》</div>""", unsafe_allow_html=True)