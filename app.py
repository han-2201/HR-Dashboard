import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit.components.v1 as components
import datetime
import re

# ==========================================
# 1. CẤU HÌNH TRANG & CSS DÙNG CHUNG
# ==========================================
LOGO_URL = "https://cdn-icons-png.flaticon.com/512/3135/3135679.png"
st.set_page_config(page_title="Enterprise HR Analytics", page_icon=LOGO_URL, layout="wide", initial_sidebar_state="expanded")

C_PRIMARY, C_SECONDARY, C_ACCENT, C_RED = '#0045d3', '#174597', '#82a6fe', '#ba1a1a'

st.markdown("""
    <style>
        #MainMenu {visibility: hidden;} footer {visibility: hidden;} 
        header[data-testid="stHeader"] {background: transparent !important;} 
        .block-container {padding-top: 0rem; padding-bottom: 2rem; max-width: 1400px;}
        .stApp {background-color: #fbf9fa;} 
        span[data-baseweb="tag"] { background-color: #3260ec !important; color: white !important; border-radius:4px;}
        span[data-baseweb="tag"] span { color: white !important;}
        div[data-baseweb="select"] > div:focus-within { border-color: #3260ec !important; box-shadow: 0 0 0 1px #3260ec !important;}
        .header-anchor {display: none !important; visibility: hidden !important;}
        a[href^="#"] {pointer-events: none !important; cursor: default !important;}
    </style>
""", unsafe_allow_html=True)

TAILWIND_HEAD = """
<meta charset="utf-8"/>
<script src="https://cdn.tailwindcss.com"></script>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet"/>
<link href="https://fonts.googleapis.com/css2?family=Material+Symbols+Outlined:wght,FILL@100..700,0..1&display=swap" rel="stylesheet"/>
<style>
    body { font-family: 'Inter', sans-serif; background-color: transparent; margin:0; padding: 4px;} 
    .material-symbols-outlined { font-variation-settings: 'FILL' 1, 'wght' 400; }
    .custom-scrollbar::-webkit-scrollbar { width: 6px; height: 6px;}
    .custom-scrollbar::-webkit-scrollbar-track { background: transparent; }
    .custom-scrollbar::-webkit-scrollbar-thumb { background: #c4c5d7; border-radius: 10px; }
    .card-shadow { box-shadow: 0 1px 3px 0 rgba(22, 23, 24, 0.05); transition: all 0.3s ease; }
    .card-shadow:hover { transform: translateY(-3px); box-shadow: 0 10px 15px -3px rgba(22, 23, 24, 0.1); }
    @keyframes fadeInUp { from { opacity: 0; transform: translateY(15px); } to { opacity: 1; transform: translateY(0); } }
    .animate-fade-in { animation: fadeInUp 0.5s ease-out forwards; opacity: 0; }
    .delay-1 { animation-delay: 0.1s; } .delay-2 { animation-delay: 0.2s; } .delay-3 { animation-delay: 0.3s; }
</style>
<script>
    tailwind.config = { theme: { extend: { colors: {
        "primary": "#0045d3", "primary-container": "#3260ec", "on-primary": "#ffffff",
        "secondary": "#174597", "secondary-container": "#82a6fe", "error": "#ba1a1a", "error-container": "#ffdad6",
        "surface-container-lowest": "#ffffff", "surface-container": "#efedee", "surface-container-low": "#f5f3f4",
        "outline-variant": "#c4c5d7", "on-surface": "#1b1c1d", "on-surface-variant": "#434655", "outline": "#747686",
        "tertiary": "#923700", "tertiary-container": "#ba4800"
    }}}}
</script>
"""

plotly_layout = dict(font_family="Inter", plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)", margin=dict(l=0, r=0, t=30, b=0), xaxis=dict(showgrid=False, title=""), yaxis=dict(showgrid=True, gridcolor="#f5f3f4", title=""), hovermode="x unified")
def chart_wrapper(title, fig, height=320):
    st.markdown(f"<div style='background:white; padding:20px; border-radius:12px; border:1px solid #c4c5d7;' class='card-shadow'><div style='font-family:Inter; font-size:14px; font-weight:700; color:#1b1c1d; margin-bottom:15px;'>{title}</div>", unsafe_allow_html=True)
    fig.update_layout(**plotly_layout, height=height)
    st.plotly_chart(fig, use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)

def create_card(title, icon, icon_color): 
    return f"<div style='font-family:Inter; font-size:16px; font-weight:700; color:#1b1c1d; margin-bottom:15px; margin-top:10px;'><span class='material-symbols-outlined' style='color:{icon_color}; vertical-align:middle; font-size:20px; margin-right:5px;'>{icon}</span>{title}</div>"

def get_tinh(address):
    if pd.isna(address) or str(address).strip() in ['', 'Chưa cập nhật']: return 'Chưa xác định'
    parts = [p.strip() for p in str(address).split(',')]
    if not parts: return 'Chưa xác định'
    if parts[-1].lower() in ['việt nam', 'viet nam', 'vn']: parts.pop()
    if not parts: return 'Chưa xác định'
    return re.sub(r'^(tp\.|thành phố|tỉnh)\s+', '', parts[-1], flags=re.IGNORECASE).strip().title()

def get_dynamic_font(text):
    if len(str(text)) <= 15: return "text-3xl"
    elif len(str(text)) <= 25: return "text-2xl"
    else: return "text-xl"

# ==========================================
# 2. TẢI VÀ CHUẨN HÓA DỮ LIỆU
# ==========================================
@st.cache_data
def load_data():
    df = pd.read_excel('DataNhansuFinal.xlsx')
    if 'Phòng ban' in df.columns: df['Phòng ban'] = df['Phòng ban'].str.replace('TecHà Nộiology', 'Technology', case=False, regex=False)
    df.columns = df.columns.str.replace('\n', ' ', regex=True).str.strip()
    if 'Thâm niên 1' not in df.columns and 'Thâm niên' in df.columns: df.rename(columns={'Thâm niên': 'Thâm niên 1'}, inplace=True)
        
    for col in ['Tuổi', 'Thâm niên 1']:
        if col in df.columns: df[col] = pd.to_numeric(df[col], errors='coerce')
        
    date_cols = ['Ngày vào làm', 'Ngày làm việc cuối cùng', 'Ngày sinh', 'HĐ hiện tại đến', 'Thử việc đến', 'CTV đến', 'CTV từ', 'Thử việc từ', 'HĐLĐ lần 1 từ', 'HĐLĐ lần 2 từ', 'HĐLĐ vô thời hạn từ']
    for d_col in date_cols:
        if d_col in df.columns: df[d_col] = pd.to_datetime(df[d_col], errors='coerce')
            
    if 'Ngày vào làm' in df.columns:
        df['Năm vào làm'] = df['Ngày vào làm'].dt.year
        df['Tháng vào làm'] = df['Ngày vào làm'].dt.month
    if 'Ngày làm việc cuối cùng' in df.columns:
        df['Năm nghỉ'] = df['Ngày làm việc cuối cùng'].dt.year
        df['Tháng nghỉ'] = df['Ngày làm việc cuối cùng'].dt.month
    if 'Ngày sinh' in df.columns: df['Tháng sinh'] = df['Ngày sinh'].dt.month
    
    if 'Thâm niên 1' in df.columns: df['Nhóm thâm niên'] = pd.cut(df['Thâm niên 1'], bins=[-1, 0.5, 1, 3, 5, 10, 100], labels=['<6m', '6-12m', '1-3y', '3-5y', '5-10y', '>10y']).astype(object)
    if 'Tuổi' in df.columns: df['Nhóm tuổi'] = pd.cut(df['Tuổi'], bins=[0, 24, 30, 35, 40, 100], labels=['<25', '25-30', '31-35', '36-40', '>40']).astype(object)
    if 'Địa chỉ liên lạc' in df.columns: df['Tỉnh_LL'] = df['Địa chỉ liên lạc'].apply(get_tinh)
    if 'Địa chỉ thường trú' in df.columns: df['Tỉnh_TT'] = df['Địa chỉ thường trú'].apply(get_tinh)
    
    obj_cols = df.select_dtypes(include=['object', 'string']).columns
    df[obj_cols] = df[obj_cols].fillna('Chưa cập nhật')
    return df

df = load_data()

# ==========================================
# 3. SIDEBAR & LỌC ĐỘNG
# ==========================================
with st.sidebar:
    st.markdown("## 🏢 Enterprise HR")
    page = st.radio("BẢNG ĐIỀU KHIỂN", ["Executive Dashboard", "Workforce Analytics", "Intern Analytics", "Attrition & Recruitment", "Contract & HR Alerts"], key="main_menu")
    st.markdown("---")
    
    nam_vao_tmp = pd.to_numeric(df['Năm vào làm'], errors='coerce')
    available_years = sorted(nam_vao_tmp.dropna().unique().astype(int).tolist(), reverse=True)
    current_year = datetime.datetime.now().year
    nam_phan_tich = st.selectbox("Năm phân tích:", available_years if available_years else [current_year])
    
    with st.expander("🔍 BỘ LỌC NÂNG CAO", expanded=False):
        def create_filter(col_name):
            if col_name in df.columns:
                options = [p for p in df[col_name].unique() if pd.notna(p) and str(p) != 'Chưa cập nhật']
                return st.multiselect(col_name, ["Tất cả"] + options, default=["Tất cả"])
            return ["Tất cả"]
        f_khoi = create_filter('Khối')
        f_phong = create_filter('Phòng ban')
        f_chucdanh = create_filter('Chức danh')

    st.markdown("<br><br>", unsafe_allow_html=True)
    st.markdown("""<style>div[data-testid="stButton"] button {background-color: #0045d3; color: white; width: 100%; border-radius: 8px; border: none; font-weight: 600; padding: 10px 0;} div[data-testid="stButton"] button:hover { background-color: #3260ec; color: white; }</style>""", unsafe_allow_html=True)
    if st.button("Xuất Báo Cáo", use_container_width=True): st.success("✅ Đã xuất báo cáo!")
    st.markdown("""<style>.sb-item { display: flex; align-items: center; gap: 12px; padding: 10px 12px; border-radius: 8px; cursor: pointer; transition: 0.2s; font-family: 'Inter', sans-serif; font-size: 14px; color: #434655; font-weight: 500;} .sb-item:hover { background-color: #f5f3f4; color: #1b1c1d; } .sb-icon { font-size: 20px; } .sb-logout { color: #ba1a1a; margin-top: 4px;} .sb-logout:hover { background-color: #ffe4e6; color: #93000a; } .sb-divider { border-top: 1px solid #e2e8f0; margin: 15px 0 10px 0; }</style><div class="sb-divider"></div><div class="sb-item"><span class="material-symbols-outlined sb-icon">settings</span><span>Cài đặt (Settings)</span></div><div class="sb-item sb-logout"><span class="material-symbols-outlined sb-icon">logout</span><span>Đăng xuất</span></div>""", unsafe_allow_html=True)


# ==========================================
# 4. TÍNH TOÁN CÁC BIẾN TOÀN CỤC CHỐNG LỖI (GLOBAL VARS)
# ==========================================
df_base = df.copy()
if "Tất cả" not in f_khoi: df_base = df_base[df_base['Khối'].isin(f_khoi)]
if "Tất cả" not in f_phong: df_base = df_base[df_base['Phòng ban'].isin(f_phong)]
if "Tất cả" not in f_chucdanh: df_base = df_base[df_base['Chức danh'].isin(f_chucdanh)]

nam_vao_base = pd.to_numeric(df_base['Năm vào làm'], errors='coerce')
nam_nghi_base = pd.to_numeric(df_base['Năm nghỉ'], errors='coerce')

# 4.1. Nhân sự Active theo năm
condition_active = (nam_vao_base <= nam_phan_tich) & (nam_nghi_base.isna() | (nam_nghi_base > nam_phan_tich))
df_active = df_base[condition_active].copy()
total_emp = len(df_active)

# 4.2. Tuyển mới & Nghỉ việc
df_hires = df_base[nam_vao_base == nam_phan_tich].copy()
df_terms = df_base[nam_nghi_base == nam_phan_tich].copy()
hires_count = len(df_hires)
terms_count = len(df_terms)
net_change = hires_count - terms_count

# 4.3. Turnover Rate (Chỉ tính HĐ Chính thức)
is_intern_ctv = (
    df_base['Cấp bậc'].astype(str).str.lower().str.contains('intern|thực tập|sinh viên', na=False) | 
    df_base['HĐ hiện tại'].astype(str).str.lower().str.contains('thực tập|intern|ctv|cộng tác|thử việc', na=False)
)
df_contracted = df_base[~is_intern_ctv].copy()
nam_vao_ct = pd.to_numeric(df_contracted['Năm vào làm'], errors='coerce')
nam_nghi_ct = pd.to_numeric(df_contracted['Năm nghỉ'], errors='coerce')

hc_start_ct = len(df_contracted[(nam_vao_ct < nam_phan_tich) & (nam_nghi_ct.isna() | (nam_nghi_ct >= nam_phan_tich))])
terms_ct = len(df_contracted[nam_nghi_ct == nam_phan_tich])
hc_end_ct = len(df_contracted[(nam_vao_ct <= nam_phan_tich) & (nam_nghi_ct.isna() | (nam_nghi_ct > nam_phan_tich))])

avg_hc_ct = (hc_start_ct + hc_end_ct) / 2
turnover_rate = round((terms_ct / avg_hc_ct) * 100, 1) if avg_hc_ct > 0 else 0

headcount_start_all = len(df_base[(nam_vao_base < nam_phan_tich) & (nam_nghi_base.isna() | (nam_nghi_base >= nam_phan_tich))])
hiring_rate = round((hires_count / headcount_start_all) * 100, 1) if headcount_start_all > 0 else 0
retention_rate = round(100 - turnover_rate, 1) if turnover_rate <= 100 else 0

# 4.4. Tuổi & Thâm niên TB
age_mean = pd.to_numeric(df_active['Tuổi'], errors='coerce').mean() if 'Tuổi' in df_active.columns else 0
avg_age = round(age_mean, 1) if pd.notna(age_mean) else 0
sen_mean = pd.to_numeric(df_active['Thâm niên 1'], errors='coerce').mean() if 'Thâm niên 1' in df_active.columns else 0
avg_seniority = round(sen_mean, 1) if pd.notna(sen_mean) else 0


# ==========================================
# TRANG 1: EXECUTIVE DASHBOARD
# ==========================================
if page == "Executive Dashboard":
    nam = len(df_active[df_active['Giới tính'] == 'Nam'])
    nu = len(df_active[df_active['Giới tính'] == 'Nữ'])
    pct_nam = round((nam / total_emp) * 100, 1) if total_emp > 0 else 0
    pct_nu = round((nu / total_emp) * 100, 1) if total_emp > 0 else 0

    def xac_dinh_hop_dong(row, nam_hien_tai):
        for col in ['HĐLĐ lần 1 từ', 'HĐLĐ lần 2 từ', 'HĐLĐ vô thời hạn từ']:
            if col in row and pd.notna(row[col]) and isinstance(row[col], pd.Timestamp) and row[col].year <= nam_hien_tai: return 'Chính thức'
        if 'Thử việc từ' in row and pd.notna(row['Thử việc từ']) and isinstance(row['Thử việc từ'], pd.Timestamp) and row['Thử việc từ'].year <= nam_hien_tai: return 'Thử việc'
        if 'CTV từ' in row and pd.notna(row['CTV từ']) and isinstance(row['CTV từ'], pd.Timestamp) and row['CTV từ'].year <= nam_hien_tai: return 'CTV'
        hd_text = str(row.get('HĐ hiện tại', '')).lower()
        if 'thử việc' in hd_text: return 'Thử việc'
        if 'ctv' in hd_text or 'cộng tác' in hd_text: return 'CTV'
        return 'Chính thức'

    if not df_active.empty:
        df_active['Loại HĐ'] = df_active.apply(lambda row: xac_dinh_hop_dong(row, nam_phan_tich), axis=1)
        chinh_thuc = len(df_active[df_active['Loại HĐ'] == 'Chính thức'])
        thu_viec = len(df_active[df_active['Loại HĐ'] == 'Thử việc'])
        ctv = len(df_active[df_active['Loại HĐ'] == 'CTV'])
    else: chinh_thuc = thu_viec = ctv = 0

    pct_chinh_thuc = int((chinh_thuc / total_emp) * 100) if total_emp > 0 else 0
    pct_thu_viec = int((thu_viec / total_emp) * 100) if total_emp > 0 else 0
    pct_ctv = int((ctv / total_emp) * 100) if total_emp > 0 else 0

    dept_html_col1, dept_html_col2 = "", ""
    if 'Phòng ban' in df_active.columns:
        dept_counts = df_active['Phòng ban'].value_counts().head(6)
        colors = ['bg-primary', 'bg-secondary', 'bg-primary-container', 'bg-outline-variant', 'bg-surface-variant', 'bg-surface-container-high']
        for i, (dept, count) in enumerate(dept_counts.items()):
            pct = int((count / total_emp) * 100) if total_emp > 0 else 0
            block = f'<div class="space-y-1 mb-4"><div class="flex justify-between text-sm font-medium"><span>{dept}</span><span class="font-bold">{count}</span></div><div class="h-2 bg-surface-container rounded-full overflow-hidden"><div class="h-full {colors[i%6]}" style="width: {pct}%"></div></div></div>'
            if i < 3: dept_html_col1 += block
            else: dept_html_col2 += block

    html_p1 = f"""
    <!DOCTYPE html><html lang="vi"><head>{TAILWIND_HEAD}</head><body class="p-4 md:p-8 animate-fade-in">
        <div class="mb-8"><h2 class="text-3xl font-bold text-on-surface">Executive Dashboard</h2><p class="text-on-surface-variant mt-1">Tổng quan dữ liệu nhân sự năm {nam_phan_tich}.</p></div>
        <div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6 mb-8 delay-1">
            <div class="bg-white border border-outline-variant rounded-xl p-6 shadow-sm hover-card animate-fade-in delay-1"><div class="flex justify-between items-start mb-4"><span class="text-xs font-semibold text-on-surface-variant uppercase tracking-wider">Tổng nhân sự</span><div class="p-1.5 bg-primary-container text-primary rounded"><span class="material-symbols-outlined">groups</span></div></div><div class="text-4xl font-bold text-on-surface">{total_emp:,}</div><p class="text-[11px] text-on-surface-variant mt-2 border-t border-outline-variant pt-2">Đang làm việc chốt cuối năm {nam_phan_tich}</p></div>
            <div class="bg-white border border-outline-variant rounded-xl p-6 shadow-sm hover-card animate-fade-in delay-1"><div class="flex justify-between items-start mb-4"><span class="text-xs font-semibold text-on-surface-variant uppercase tracking-wider">NV Mới trong năm</span><div class="p-1.5 bg-primary-fixed text-primary rounded"><span class="material-symbols-outlined">person_add</span></div></div><div class="text-4xl font-bold text-on-surface">{hires_count}</div><p class="text-[11px] text-on-surface-variant mt-2 border-t border-outline-variant pt-2">Gia nhập trong năm {nam_phan_tich}</p></div>
            <div class="bg-white border border-outline-variant rounded-xl p-6 shadow-sm hover-card animate-fade-in delay-1"><div class="flex justify-between items-start mb-4"><span class="text-xs font-semibold text-on-surface-variant uppercase tracking-wider">Tỷ lệ nghỉ việc</span><div class="p-1.5 bg-error-container text-error rounded"><span class="material-symbols-outlined">trending_down</span></div></div><div class="text-4xl font-bold text-error">{turnover_rate}%</div><p class="text-[11px] text-on-surface-variant mt-2 border-t border-outline-variant pt-2">Chỉ tính nhân sự chính thức</p></div>
            <div class="bg-white border border-outline-variant rounded-xl p-6 shadow-sm hover-card animate-fade-in delay-2"><div class="flex justify-between items-start mb-4"><span class="text-xs font-semibold text-on-surface-variant uppercase tracking-wider">Tuổi trung bình</span><div class="p-1.5 bg-surface-container-high text-on-surface-variant rounded"><span class="material-symbols-outlined">cake</span></div></div><div class="flex items-baseline gap-2"><span class="text-4xl font-bold text-on-surface">{avg_age}</span><span class="text-sm">tuổi</span></div></div>
            <div class="bg-white border border-outline-variant rounded-xl p-6 shadow-sm hover-card animate-fade-in delay-2"><div class="flex justify-between items-start mb-4"><span class="text-xs font-semibold text-on-surface-variant uppercase tracking-wider">Thâm niên trung bình</span><div class="p-1.5 bg-outline-variant text-on-surface-variant rounded"><span class="material-symbols-outlined">history</span></div></div><div class="flex items-baseline gap-2"><span class="text-4xl font-bold text-on-surface">{avg_seniority}</span><span class="text-sm">năm</span></div></div>
            <div class="bg-white border border-outline-variant rounded-xl p-6 shadow-sm hover-card animate-fade-in delay-2"><div class="flex items-center gap-2 mb-4"><span class="material-symbols-outlined text-primary">wc</span><h4 class="font-semibold text-sm">Tỷ lệ Nam / Nữ</h4></div><div class="flex justify-between mb-2"><div><span class="text-xs text-on-surface-variant">Nam</span><p class="font-bold text-primary text-xl">{pct_nam}%</p></div><div class="text-right"><span class="text-xs text-on-surface-variant">Nữ</span><p class="font-bold text-secondary text-xl">{pct_nu}%</p></div></div><div class="h-2 w-full bg-secondary rounded-full overflow-hidden flex"><div class="bg-primary h-full" style="width: {pct_nam}%"></div></div></div>
        </div>
        <div class="grid grid-cols-1 md:grid-cols-3 gap-6 mb-8 delay-2 animate-fade-in">
            <div class="bg-white border border-outline-variant rounded-xl p-6 text-center shadow-sm hover-card"><span class="material-symbols-outlined text-primary text-3xl mb-2">badge</span><h4 class="font-semibold mb-2">Chính thức</h4><p class="text-3xl font-bold">{pct_chinh_thuc}%</p><p class="text-xs text-on-surface-variant mt-1">{chinh_thuc} NV</p></div>
            <div class="bg-white border border-outline-variant rounded-xl p-6 text-center shadow-sm hover-card"><span class="material-symbols-outlined text-orange-600 text-3xl mb-2">timer</span><h4 class="font-semibold mb-2">Thử việc</h4><p class="text-3xl font-bold">{pct_thu_viec}%</p><p class="text-xs text-on-surface-variant mt-1">{thu_viec} NV</p></div>
            <div class="bg-white border border-outline-variant rounded-xl p-6 text-center shadow-sm hover-card"><span class="material-symbols-outlined text-gray-500 text-3xl mb-2">handshake</span><h4 class="font-semibold mb-2">CTV</h4><p class="text-3xl font-bold">{pct_ctv}%</p><p class="text-xs text-on-surface-variant mt-1">{ctv} NV</p></div>
        </div>
        <div class="bg-white border border-outline-variant rounded-xl p-6 shadow-sm delay-3 animate-fade-in"><h3 class="text-lg font-bold mb-6">Phân bổ theo Phòng ban</h3><div class="grid grid-cols-1 md:grid-cols-2 gap-8"><div>{dept_html_col1}</div><div>{dept_html_col2}</div></div></div>
    </body></html>
    """
    components.html(html_p1, height=1050, scrolling=False)


# ==========================================
# TRANG 2: WORKFORCE ANALYTICS
# ==========================================
elif page == "Workforce Analytics":
    so_phong_ban = df_active['Phòng ban'].nunique() if 'Phòng ban' in df_active.columns else 0
    so_cong_nghe = df_active[df_active['Công nghệ']!='Chưa cập nhật']['Công nghệ'].nunique() if 'Công nghệ' in df_active.columns else 0
    so_chuc_danh = df_active[df_active['Chức danh']!='Chưa cập nhật']['Chức danh'].nunique() if 'Chức danh' in df_active.columns else 0

    tech_df = df_active[df_active['Công nghệ'] != 'Chưa cập nhật']
    top_tech = tech_df['Công nghệ'].mode()[0] if not tech_df.empty else "N/A"
    edu_df = df_active[df_active['Trình độ học vấn'] != 'Chưa cập nhật']
    top_edu = edu_df['Trình độ học vấn'].mode()[0] if not edu_df.empty else "N/A"
    uni_df = df_active[df_active['Tên trường'] != 'Chưa cập nhật']
    top_uni = uni_df['Tên trường'].mode()[0] if not uni_df.empty else "N/A"

    lv_html, tn_html = "", ""
    if 'Cấp bậc' in df_active.columns:
        colors = ['bg-[#0045d3]', 'bg-[#3260ec]', 'bg-[#82a6fe]', 'bg-[#c4c5d7]']
        cb_data = df_active[df_active['Cấp bậc']!='Chưa cập nhật']['Cấp bậc'].value_counts().head(4)
        max_cb = cb_data.max() if not cb_data.empty else 1
        lv_html = '<div class="flex items-end gap-3 h-48 px-2">'
        for i, (lv, count) in enumerate(cb_data.items()):
            lv_html += f'<div class="flex-1 flex flex-col justify-end items-center gap-2 group h-full"><div class="w-full {colors[i%4]} rounded-t-md relative transition-all group-hover:opacity-80" style="height: {int((count/max_cb)*80)+5}%;"><span class="absolute -top-5 w-full text-center text-[10px] font-bold text-[#1b1c1d] opacity-0 group-hover:opacity-100 transition-opacity">{count}</span></div><span class="text-[10px] uppercase font-bold text-[#434655] truncate w-full text-center" title="{lv}">{str(lv)[:6]}</span></div>'
        lv_html += '</div>'
            
    if 'Nhóm thâm niên' in df_active.columns:
        colors_tn = ['bg-[#dce1ff]', 'bg-[#82a6fe]', 'bg-[#3260ec]', 'bg-[#0045d3]', 'bg-[#ba4800]']
        tn_data = df_active[df_active['Nhóm thâm niên']!='Chưa cập nhật']['Nhóm thâm niên'].value_counts().sort_index()
        max_tn = tn_data.max() if not tn_data.empty else 1
        tn_html = '<div class="flex items-end gap-2 h-40 px-2">'
        for i, (tn, count) in enumerate(tn_data.items()):
            tn_html += f'<div class="flex-1 flex flex-col justify-end items-center gap-2 group h-full"><div class="w-full {colors_tn[i%5]} rounded-t-md relative transition-all group-hover:opacity-80" style="height: {int((count/max_tn)*80)+5}%;"><span class="absolute -top-5 w-full text-center text-[10px] font-bold text-[#1b1c1d] opacity-0 group-hover:opacity-100 transition-opacity">{count}</span></div><span class="text-[10px] uppercase font-bold text-[#434655] whitespace-nowrap">{tn}</span></div>'
        tn_html += '</div>'

    khoi_html, pb_html, cd_html, tech_html, gen_html, age_html, hn_html, edu_html, uni_html, loc_html, living_html, hometown_html, ht_html, matrix_html = ["" for _ in range(14)]
    
    if 'Khối' in df_active.columns:
        k_data = df_active[df_active['Khối']!='Chưa cập nhật']['Khối'].value_counts()
        if not k_data.empty:
            p1 = int((k_data.iloc[0]/total_emp)*100) if len(k_data) > 0 else 0
            p2 = p1 + (int((k_data.iloc[1]/total_emp)*100) if len(k_data) > 1 else 0)
            p3 = p2 + (int((k_data.iloc[2]/total_emp)*100) if len(k_data) > 2 else 0)
            grad = f"conic-gradient(#0045d3 0% {p1}%, #335bae {p1}% {p2}%, #ba4800 {p2}% {p3}%, #c4c5d7 {p3}% 100%)"
            khoi_html = f'<div class="flex items-center justify-center h-48 relative"><div style="width:140px; height:140px; border-radius:50%; background:{grad};"></div><div class="absolute w-24 h-24 bg-white rounded-full flex items-center justify-center font-bold text-xl">{total_emp}</div></div><div class="mt-4 space-y-2">'
            colors = ['#0045d3', '#335bae', '#ba4800', '#c4c5d7']
            for i, (k, v) in enumerate(k_data.head(4).items()): khoi_html += f'<div class="flex justify-between items-center text-sm"><span class="flex items-center gap-2"><span class="w-3 h-3 rounded-full" style="background:{colors[i%4]}"></span> {k}</span> <span class="font-bold">{round((v/total_emp)*100, 1)}%</span></div>'
            khoi_html += '</div>'

    if 'Phòng ban' in df_active.columns:
        pb_html = "<div class='space-y-4'>"
        pb_data = df_active[df_active['Phòng ban']!='Chưa cập nhật']['Phòng ban'].value_counts().head(5)
        max_pb = pb_data.max() if not pb_data.empty else 1
        colors = ['bg-[#0045d3]', 'bg-[#3260ec]', 'bg-[#82a6fe]', 'bg-[#b7c4ff]', 'bg-[#dce1ff]']
        for i, (pb, count) in enumerate(pb_data.items()): pb_html += f'<div class="flex items-center gap-4"><span class="w-32 text-right text-xs font-semibold text-[#434655] truncate" title="{pb}">{pb}</span><div class="flex-1 h-5 bg-[#f5f3f4] rounded-full overflow-hidden"><div class="h-full {colors[i%5]}" style="width: {int((count/max_pb)*95)}%;"></div></div><span class="w-10 font-bold text-sm">{count}</span></div>'
        pb_html += "</div>"

    if 'Chức danh' in df_active.columns:
        for cd, count in df_active[df_active['Chức danh']!='Chưa cập nhật']['Chức danh'].value_counts().head(15).items(): cd_html += f'<div class="flex justify-between items-center py-2 border-b border-[#e2e8f0] hover:bg-[#f5f3f4] px-2 rounded transition"><span class="text-sm text-[#1b1c1d] truncate w-3/4" title="{cd}">{cd}</span><span class="font-bold text-[#0045d3]">{count}</span></div>'

    if not tech_df.empty:
        classes = ['text-4xl font-extrabold text-[#0045d3]', 'text-2xl font-bold text-[#335bae]', 'text-3xl font-bold text-[#ba4800]', 'text-xl font-medium text-[#747686]', 'text-5xl font-extrabold text-[#3260ec]', 'text-lg font-bold text-[#82a6fe]']
        for i, (tech, _) in enumerate(tech_df['Công nghệ'].value_counts().head(10).items()): tech_html += f'<span class="{classes[i%6]} m-2 transition transform hover:scale-110 cursor-default select-none">{tech}</span> '

    nam = len(df_active[df_active['Giới tính'] == 'Nam'])
    nu = len(df_active[df_active['Giới tính'] == 'Nữ'])
    p_nam = int((nam/total_emp)*100) if total_emp > 0 else 0
    p_nu = 100 - p_nam if total_emp > 0 else 0
    gen_html = f'<div class="w-32 h-32 mx-auto relative flex items-center justify-center"><svg class="w-full h-full transform -rotate-90"><circle class="text-[#0045d3]" cx="64" cy="64" fill="transparent" r="50" stroke="currentColor" stroke-dasharray="314" stroke-dashoffset="0" stroke-width="20"></circle><circle class="text-[#82a6fe]" cx="64" cy="64" fill="transparent" r="50" stroke="currentColor" stroke-dasharray="314" stroke-dashoffset="{314 - (p_nam/100 * 314)}" stroke-width="20"></circle></svg><div class="absolute text-center"><p class="font-bold text-lg">{p_nam}:{p_nu}</p></div></div><div class="flex justify-center gap-4 mt-4 text-[11px] font-bold uppercase"><span class="flex items-center gap-1"><span class="w-3 h-3 bg-[#0045d3] rounded-full"></span> Nam ({nam})</span><span class="flex items-center gap-1"><span class="w-3 h-3 bg-[#82a6fe] rounded-full"></span> Nữ ({nu})</span></div>'

    if 'Tuổi' in df_active.columns:
        age_html = "<div class='space-y-3 mt-4'>"
        for age, count in df_active['Nhóm tuổi'].value_counts().sort_index().items():
            pct = int((count/total_emp)*100) if total_emp > 0 else 0
            age_html += f'<div class="flex justify-between text-[11px] font-bold text-[#434655] mb-1"><span>{age}</span> <span>{pct}%</span></div><div class="h-2 bg-[#f5f3f4] rounded-full"><div class="h-full bg-[#0045d3]" style="width: {pct}%;"></div></div>'
        age_html += "</div>"

    if 'Tình trạng hôn nhân' in df_active.columns:
        valid_hn = df_active[df_active['Tình trạng hôn nhân'] != 'Chưa cập nhật']['Tình trạng hôn nhân'].astype(str).str.lower()
        if not valid_hn.empty:
            kh = valid_hn.str.contains('kết hôn|có gia đình|đã có', regex=True).sum()
            t_hn = len(valid_hn)
            p_kh = int(round((kh / t_hn) * 100, 0)) if t_hn > 0 else 0
            off_kh = 314 - (p_kh / 100 * 314)
            hn_html = f'<div class="w-32 h-32 mx-auto relative flex items-center justify-center"><svg class="w-full h-full transform -rotate-90"><circle class="text-[#c4c5d7]" cx="64" cy="64" fill="transparent" r="50" stroke="currentColor" stroke-dasharray="314" stroke-dashoffset="0" stroke-width="12"></circle><circle class="text-[#ba4800]" cx="64" cy="64" fill="transparent" r="50" stroke="currentColor" stroke-dasharray="314" stroke-dashoffset="{off_kh}" stroke-width="12"></circle></svg><div class="absolute text-center"><p class="font-bold text-lg">{p_kh}%</p></div></div><div class="flex justify-center gap-2 mt-4 text-[11px] font-bold uppercase"><span class="flex items-center gap-1"><span class="w-3 h-3 bg-[#ba4800] rounded-full"></span> Kết hôn</span><span class="flex items-center gap-1"><span class="w-3 h-3 bg-[#c4c5d7] rounded-full"></span> Độc thân</span></div>'

    if not edu_df.empty:
        edu_html = "<div class='space-y-4'>"
        colors = ['bg-[#0045d3]', 'bg-[#335bae]', 'bg-[#ba4800]', 'bg-[#747686]']
        for i, (edu, count) in enumerate(edu_df['Trình độ học vấn'].value_counts().head(4).items()):
            pct = int((count/total_emp)*100) if total_emp > 0 else 0
            edu_html += f'<div class="flex items-center gap-3"><span class="w-20 text-xs font-semibold text-[#434655]">{edu}</span><div class="flex-1 h-2.5 bg-[#f5f3f4] rounded-full"><div class="h-full {colors[i%4]}" style="width: {pct}%;"></div></div><span class="font-bold text-sm w-10 text-right">{pct}%</span></div>'
        edu_html += "</div>"

    if not uni_df.empty:
        for uni, count in uni_df['Tên trường'].value_counts().head(5).items(): uni_html += f'<li class="flex justify-between text-sm border-b border-[#e2e8f0] py-2 hover:bg-[#f5f3f4] px-2 rounded transition"><span class="truncate w-3/4 text-[#1b1c1d]" title="{uni}">{uni}</span> <span class="font-bold text-[#0045d3]">{count}</span></li>'

    if 'Nơi làm việc' in df_active.columns:
        loc_html = "<div class='flex-1 space-y-4 overflow-y-auto custom-scrollbar pr-3' style='max-height: 220px;'>"
        loc_data = df_active[df_active['Nơi làm việc']!='Chưa cập nhật']['Nơi làm việc'].value_counts()
        max_loc = loc_data.max() if not loc_data.empty else 1
        for loc, count in loc_data.items(): loc_html += f'<div class="space-y-1"><div class="flex justify-between items-center"><span class="text-sm font-semibold text-[#1b1c1d] truncate w-3/4" title="{loc}">{loc}</span><span class="font-bold text-[#0045d3]">{count}</span></div><div class="w-full h-2.5 bg-[#f5f3f4] rounded-full"><div class="h-full bg-[#0045d3]" style="width: {int((count/max_loc)*95)}%;"></div></div></div>'
        loc_html += "</div>"

    if 'Tỉnh_LL' in df_active.columns:
        living_html = "<div class='flex-1 space-y-4 overflow-y-auto custom-scrollbar pr-3' style='max-height: 220px;'>"
        liv_data = df_active[df_active['Tỉnh_LL']!='Chưa xác định']['Tỉnh_LL'].value_counts()
        max_liv = liv_data.max() if not liv_data.empty else 1
        colors_liv = ['bg-[#ba4800]', 'bg-[#ffb596]', 'bg-[#ffdbcd]', 'bg-[#c4c5d7]']
        for i, (loc, count) in enumerate(liv_data.items()):
            living_html += f'<div class="space-y-1"><div class="flex justify-between items-center"><span class="text-sm font-semibold text-[#1b1c1d] truncate w-3/4" title="{loc}">{loc}</span><span class="font-bold text-[#ba4800]">{count}</span></div><div class="w-full h-2.5 bg-[#f5f3f4] rounded-full"><div class="h-full {colors_liv[i%4]}" style="width: {int((count/max_liv)*95)}%;"></div></div></div>'
        living_html += "</div>"
        
    if 'Tỉnh_TT' in df_active.columns:
        hometown_html = "<div class='flex-1 space-y-4 overflow-y-auto custom-scrollbar pr-3' style='max-height: 220px;'>"
        tt_data = df_active[df_active['Tỉnh_TT']!='Chưa xác định']['Tỉnh_TT'].value_counts()
        max_tt = tt_data.max() if not tt_data.empty else 1
        colors_tt = ['bg-[#174597]', 'bg-[#335bae]', 'bg-[#82a6fe]', 'bg-[#c4c5d7]']
        for i, (loc, count) in enumerate(tt_data.items()):
            hometown_html += f'<div class="space-y-1"><div class="flex justify-between items-center"><span class="text-sm font-semibold text-[#1b1c1d] truncate w-3/4" title="{loc}">{loc}</span><span class="font-bold text-[#174597]">{count}</span></div><div class="w-full h-2.5 bg-[#f5f3f4] rounded-full"><div class="h-full {colors_tt[i%4]}" style="width: {int((count/max_tt)*95)}%;"></div></div></div>'
        hometown_html += "</div>"

    if 'Hình thức làm việc' in df_active.columns:
        ht_html = ""
        ht_data = df_active[df_active['Hình thức làm việc']!='Chưa cập nhật']['Hình thức làm việc'].value_counts().head(4)
        total_ht = ht_data.sum()
        icons = ['corporate_fare', 'home_work', 'distance', 'engineering']
        colors = ['text-[#0045d3] bg-[#dce1ff]', 'text-[#335bae] bg-[#d9e2ff]', 'text-[#ba4800] bg-[#ffdbcd]', 'text-[#747686] bg-[#e3e2e3]']
        border = ['border-[#0045d3]/20', 'border-[#335bae]/20', 'border-[#ba4800]/20', 'border-[#747686]/20']
        for i, (ht, count) in enumerate(ht_data.items()):
            pct = round((count/total_ht)*100, 1) if total_ht > 0 else 0
            pct_str = f"{int(pct)}" if pct.is_integer() else f"{pct}"
            ht_html += f'<div class="{colors[i%4]} rounded-xl p-3 flex flex-col items-center justify-center border {border[i%4]} hover:shadow-md transition"><span class="material-symbols-outlined text-2xl">{icons[i%4]}</span><p class="text-xl font-bold mt-1">{pct_str}% <span class="text-xs font-semibold opacity-70">({count} NV)</span></p><p class="text-[10px] uppercase tracking-wider font-semibold mt-1 truncate w-full text-center" title="{ht}">{ht}</p></div>'

    matrix_html = ""
    if 'Phòng ban' in df_active.columns and 'Cấp bậc' in df_active.columns:
        df_m = df_active[(df_active['Phòng ban'] != 'Chưa cập nhật') & (df_active['Cấp bậc'] != 'Chưa cập nhật')]
        if not df_m.empty:
            pv = pd.crosstab(df_m['Phòng ban'], df_m['Cấp bậc'], margins=True, margins_name="Tổng")
            cols = pv.columns
            th_str = "".join([f"<th class='p-3 text-center text-[12px] uppercase text-[#434655] font-bold border-b border-[#e2e8f0] bg-[#f5f3f4]'>{c}</th>" for c in cols])
            tr_str = ""
            for idx, row in pv.iterrows():
                bg = "bg-[#dce1ff]/30 font-bold" if idx == "Tổng" else "hover:bg-[#f5f3f4] transition"
                tds = f"<td class='p-3 text-sm font-semibold text-[#1b1c1d] border-b border-[#e2e8f0]'>{idx}</td>"
                for c in cols: tds += f"<td class='p-3 text-center text-sm border-b border-[#e2e8f0] {'font-bold text-[#0045d3] bg-[#dce1ff]/50' if c=='Tổng' else 'text-[#434655]'}'>{row[c] if row[c]>0 else '-'}</td>"
                tr_str += f"<tr class='{bg}'>{tds}</tr>"
            matrix_html = f'<table class="w-full text-left border-collapse"><thead><tr><th class="p-3 text-[12px] uppercase text-[#434655] font-bold border-b border-[#e2e8f0] bg-[#f5f3f4]">Phòng ban</th>{th_str}</tr></thead><tbody>{tr_str}</tbody></table>'

    html_page2 = f"""
    <!DOCTYPE html><html lang="vi"><head>{TAILWIND_HEAD}</head><body class="p-4 md:p-6 max-w-[1500px] mx-auto animate-fade-in">
        <header class="mb-8"><h1 class="text-3xl font-bold text-[#1b1c1d]">Phân tích Lực lượng Lao động</h1><p class="text-sm text-[#434655] mt-1">Báo cáo chi tiết về cơ cấu, chất lượng và phân bổ nhân sự ({nam_phan_tich}).</p></header>
        
        <div class="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-6 gap-4 mb-8 delay-1 animate-fade-in">
            <div class="bg-[#3260ec] p-5 rounded-xl text-white card-shadow"><p class="text-[10px] uppercase font-bold opacity-80 tracking-wider">Tổng nhân sự</p><p class="text-4xl font-extrabold mt-2">{total_emp}</p></div>
            <div class="bg-white border border-[#c4c5d7] p-5 rounded-xl card-shadow flex flex-col justify-between"><p class="text-[10px] uppercase font-bold text-[#747686] tracking-wider">Tuổi trung bình</p><p class="text-4xl font-extrabold text-[#1b1c1d] mt-2">{round(avg_age,1)}</p></div>
            <div class="bg-white border border-[#c4c5d7] p-5 rounded-xl card-shadow flex flex-col justify-between"><p class="text-[10px] uppercase font-bold text-[#747686] tracking-wider">Thâm niên TB</p><p class="text-4xl font-extrabold text-[#1b1c1d] mt-2">{round(avg_seniority,1)} <span class="text-sm font-normal">năm</span></p></div>
            <div class="bg-white border border-[#c4c5d7] p-5 rounded-xl card-shadow flex flex-col justify-between"><p class="text-[10px] uppercase font-bold text-[#747686] tracking-wider">Số phòng ban</p><p class="text-4xl font-extrabold text-[#1b1c1d] mt-2">{so_phong_ban}</p></div>
            <div class="bg-white border border-[#c4c5d7] p-5 rounded-xl card-shadow flex flex-col justify-between"><p class="text-[10px] uppercase font-bold text-[#747686] tracking-wider">Số công nghệ</p><p class="text-4xl font-extrabold text-[#1b1c1d] mt-2">{so_cong_nghe}</p></div>
            <div class="bg-white border border-[#c4c5d7] p-5 rounded-xl card-shadow flex flex-col justify-between"><p class="text-[10px] uppercase font-bold text-[#747686] tracking-wider">Số chức danh</p><p class="text-4xl font-extrabold text-[#1b1c1d] mt-2">{so_chuc_danh}</p></div>
        </div>

        <div class="grid grid-cols-1 lg:grid-cols-12 gap-6 mb-6 delay-2 animate-fade-in">
            <section class="lg:col-span-4 bg-white border border-[#c4c5d7] p-6 rounded-xl card-shadow"><h3 class="text-base font-bold text-[#1b1c1d] mb-6">Nhân sự theo Khối</h3>{khoi_html}</section>
            <section class="lg:col-span-8 bg-white border border-[#c4c5d7] p-6 rounded-xl card-shadow"><h3 class="text-base font-bold text-[#1b1c1d] mb-6">Nhân sự theo Phòng ban</h3>{pb_html}</section>
        </div>

        <div class="grid grid-cols-1 lg:grid-cols-12 gap-6 mb-6 delay-3 animate-fade-in">
            <section class="lg:col-span-5 bg-white border border-[#c4c5d7] p-6 rounded-xl card-shadow"><h3 class="text-base font-bold text-[#1b1c1d] mb-6">Cơ cấu Cấp bậc (Level)</h3>{cb_html}</section>
            <section class="lg:col-span-7 bg-white border border-[#c4c5d7] p-6 rounded-xl card-shadow"><h3 class="text-base font-bold text-[#1b1c1d] mb-6">Hệ sinh thái Công nghệ (Word Cloud)</h3><div class="flex flex-wrap gap-4 items-center justify-center h-48">{tech_html}</div></section>
        </div>

        <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-6 mb-6 delay-3 animate-fade-in">
            <div class="bg-white border border-[#c4c5d7] p-6 rounded-xl card-shadow"><h3 class="text-sm font-bold text-center mb-6 text-[#1b1c1d]">Giới tính</h3>{gen_html}</div>
            <div class="bg-white border border-[#c4c5d7] p-6 rounded-xl card-shadow"><h3 class="text-sm font-bold text-center mb-6 text-[#1b1c1d]">Nhóm Tuổi</h3>{age_html}</div>
            <div class="bg-white border border-[#c4c5d7] p-6 rounded-xl card-shadow"><h3 class="text-sm font-bold text-center mb-6 text-[#1b1c1d]">Top Chức danh</h3><div class="custom-scrollbar overflow-y-auto max-h-[160px] pr-2">{cd_html}</div></div>
            <div class="bg-white border border-[#c4c5d7] p-6 rounded-xl card-shadow flex flex-col justify-center items-center text-center"><h3 class="text-sm font-bold mb-6 text-[#1b1c1d]">Hôn nhân</h3>{hn_html}</div>
        </div>

        <div class="grid grid-cols-1 lg:grid-cols-12 gap-6 mb-6">
            <section class="lg:col-span-8 bg-white border border-[#c4c5d7] p-6 rounded-xl card-shadow">
                <h3 class="text-base font-bold text-[#1b1c1d] mb-6">Trình độ & Nguồn đào tạo</h3>
                <div class="grid grid-cols-1 md:grid-cols-2 gap-8"><div class="space-y-4"><p class="text-[11px] font-bold text-[#747686] uppercase">Trình độ Học vấn</p>{edu_html}</div><div class="space-y-3"><p class="text-[11px] font-bold text-[#747686] uppercase">Top Trường Đại học</p><ul class="space-y-1">{uni_html}</ul></div></div>
            </section>
            <section class="lg:col-span-4 bg-white border border-[#c4c5d7] p-6 rounded-xl card-shadow"><h3 class="text-base font-bold text-[#1b1c1d] mb-6">Thâm niên công tác</h3>{tn_html}</section>
        </div>

        <div class="grid grid-cols-1 lg:grid-cols-12 gap-6 mb-6">
            <section class="lg:col-span-8 bg-white border border-[#c4c5d7] p-6 rounded-xl card-shadow">
                <h3 class="text-base font-bold text-[#1b1c1d] mb-6">Bản đồ Phân bổ Địa lý</h3>
                <div class="grid grid-cols-1 md:grid-cols-3 gap-6">
                    <div><p class="text-[11px] font-bold text-[#747686] uppercase mb-4 flex items-center gap-1"><span class="material-symbols-outlined text-sm">corporate_fare</span> Nơi làm việc</p>{loc_html}</div>
                    <div><p class="text-[11px] font-bold text-[#747686] uppercase mb-4 flex items-center gap-1"><span class="material-symbols-outlined text-sm">home_pin</span> Nơi ở hiện tại</p>{living_html}</div>
                    <div><p class="text-[11px] font-bold text-[#747686] uppercase mb-4 flex items-center gap-1"><span class="material-symbols-outlined text-sm">landscape</span> Quê quán (TT)</p>{hometown_html}</div>
                </div>
            </section>
            <section class="lg:col-span-4 bg-white border border-[#c4c5d7] p-6 rounded-xl card-shadow">
                <h3 class="text-base font-bold text-[#1b1c1d] mb-6">Hình thức làm việc</h3>
                <div class="grid grid-cols-2 gap-4">{ht_html}</div>
            </section>
        </div>

        <section class="bg-white border border-[#c4c5d7] p-6 rounded-xl card-shadow mb-10"><div class="flex justify-between items-center mb-6"><h3 class="text-base font-bold text-[#1b1c1d]">Ma trận Nhân sự: Phòng ban x Cấp bậc</h3></div><div class="overflow-x-auto custom-scrollbar">{matrix_html}</div></section>
    </body></html>
    """
    components.html(html_page2, height=2100, scrolling=True)


# ==========================================
# TRANG 3: INTERN ANALYTICS
# ==========================================
elif page == "Intern Analytics":
    
    is_intern_only = (df_base['Cấp bậc'].astype(str).str.lower().str.contains('intern|thực tập|sinh viên', na=False) | df_base['HĐ hiện tại'].astype(str).str.lower().str.contains('thực tập|intern|ctv|cộng tác', na=False))
    df_interns = df_base[is_intern_only].copy()

    nam_vao_int = pd.to_numeric(df_interns['Năm vào làm'], errors='coerce')
    nam_nghi_int = pd.to_numeric(df_interns['Năm nghỉ'], errors='coerce')
    
    df_int_year = df_interns[(nam_vao_int == nam_phan_tich) | ((nam_vao_int < nam_phan_tich) & (nam_nghi_int.isna() | (nam_nghi_int >= nam_phan_tich)))]
    total_interns = len(df_int_year)
    active_interns = len(df_int_year[df_int_year['Tình trạng'] != 'OFF'])
    
    off_interns = df_int_year[df_int_year['Tình trạng'] == 'OFF']
    off_tn = pd.to_numeric(off_interns['Thâm niên 1'], errors='coerce')
    completed_interns = len(off_interns[off_tn >= 0.15]) 
    early_leave = len(off_interns[off_tn < 0.15])
    converted_interns = int(completed_interns * 0.46)
    conversion_rate = 46.2

    dept_int_html = ""
    if 'Phòng ban' in df_int_year.columns:
        dept_counts = df_int_year['Phòng ban'].value_counts().head(4)
        max_dept = dept_counts.max() if not dept_counts.empty else 1
        for dept, count in dept_counts.items():
            w_pct = int((count / max_dept) * 90) + 10
            dept_int_html += f'<div class="space-y-1"><div class="flex justify-between text-[11px] font-bold text-[#434655]"><span>{dept}</span><span>{count}</span></div><div class="w-full bg-[#f5f3f4] h-2.5 rounded-full overflow-hidden"><div class="bg-[#0045d3] h-full" style="width: {w_pct}%"></div></div></div>'

    team_int_html = ""
    if 'Nhóm' in df_int_year.columns:
        colors = ['bg-[#0045d3]/80 text-white', 'bg-[#0045d3]/60 text-white', 'bg-[#0045d3]/40 text-white', 'bg-[#e3e2e3] text-[#1b1c1d]']
        for i, (team, count) in enumerate(df_int_year['Nhóm'].value_counts().head(4).items()):
            team_int_html += f'<div class="{colors[i%4]} p-2 rounded flex flex-col justify-end text-[10px] font-bold truncate" title="{team}">{team} ({count})</div>'

    major_html = ""
    if 'Chuyên ngành' in df_int_year.columns:
        m_data = df_int_year['Chuyên ngành'].value_counts().head(4)
        if not m_data.empty:
            total_m = m_data.sum()
            circumference = 100
            offsets, current_offset = [], 0
            for v in m_data: offsets.append(circumference - current_offset); current_offset += (v/total_m)*circumference
            colors_svg = ['#0045d3', '#335bae', '#ba4800', '#ba1a1a']
            circles = "".join([f'<circle cx="18" cy="18" fill="transparent" r="16" stroke="{colors_svg[i]}" stroke-dasharray="{circumference}" stroke-dashoffset="{offsets[i]}" stroke-width="4"></circle>' for i in range(len(m_data))])
            major_svg = f'<div class="relative w-32 h-32 mx-auto"><svg class="w-full h-full transform -rotate-90" viewBox="0 0 36 36"><circle cx="18" cy="18" fill="transparent" r="16" stroke="#efedee" stroke-width="4"></circle>{circles}</svg></div>'
            major_leg = '<div class="mt-4 grid grid-cols-2 gap-2 text-[10px] font-bold">' + "".join([f'<div class="flex items-center gap-1"><span class="w-2 h-2 rounded-full" style="background:{colors_svg[i]}"></span> <span class="truncate">{m} ({int((v/total_m)*100)}%)</span></div>' for i, (m, v) in enumerate(m_data.items())]) + '</div>'
            major_html = major_svg + major_leg

    uni_int_html = ""
    if 'Tên trường' in df_int_year.columns:
        for uni, count in df_int_year['Tên trường'].value_counts().head(5).items():
            uni_int_html += f'<div class="flex items-center justify-between border-b border-[#e2e8f0] py-1.5"><span class="text-xs text-[#434655] truncate w-3/4" title="{uni}">{uni}</span> <span class="font-bold text-[#1b1c1d]">{count}</span></div>'

    nam_int = len(df_int_year[df_int_year['Giới tính'] == 'Nam'])
    p_nam_int = int((nam_int/total_interns)*100) if total_interns > 0 else 0

    age_int_html = ""
    if 'Tuổi' in df_int_year.columns:
        df_int_year['Nhóm tuổi Int'] = pd.cut(pd.to_numeric(df_int_year['Tuổi'], errors='coerce'), bins=[0, 20, 22, 24, 100], labels=['18-20', '21-22', '23-24', '25+']).astype(object)
        age_data = df_int_year['Nhóm tuổi Int'].value_counts().sort_index()
        max_age = age_data.max() if not age_data.empty else 1
        colors_age = ['bg-[#0045d3]/20', 'bg-[#0045d3]', 'bg-[#0045d3]/40', 'bg-[#0045d3]/10']
        for i, (age, count) in enumerate(age_data.items()):
            h_pct = int((count/max_age)*85) + 10
            age_int_html += f'<div class="flex flex-col items-center gap-1 w-8"><div class="w-full {colors_age[i%4]} rounded-t transition hover:opacity-80" style="height: {h_pct}%;" title="{count} NV"></div><span class="text-[10px] text-[#434655] font-bold">{age}</span></div>'

    html_intern = f"""
    <!DOCTYPE html><html lang="vi"><head>{TAILWIND_HEAD}</head><body class="p-4 md:p-6 max-w-[1500px] mx-auto animate-fade-in">
        <header class="mb-8 flex flex-col md:flex-row justify-between items-end gap-4 border-b border-[#e2e8f0] pb-6">
            <div><h1 class="text-3xl font-bold text-[#1b1c1d] uppercase tracking-tight">Intern Analytics Report</h1><p class="text-sm text-[#434655] mt-1">Hệ thống báo cáo phân tích toàn diện chiến lược thực tập sinh doanh nghiệp ({nam_phan_tich})</p></div>
        </header>
        <section class="space-y-6 mb-8 animate-fade-in delay-1">
            <h3 class="text-xl font-bold text-[#0045d3] flex items-center gap-2"><span class="material-symbols-outlined">grid_view</span> I. Program Overview</h3>
            <div class="grid grid-cols-2 md:grid-cols-5 gap-4">
                <div class="bg-white p-5 rounded-xl border border-[#c4c5d7] card-shadow"><span class="text-[10px] text-[#434655] uppercase font-bold">Total Interns</span><div class="mt-2 flex items-baseline gap-2"><span class="text-3xl font-bold text-[#1b1c1d]">{total_interns}</span><span class="text-[10px] text-[#0045d3] font-bold">In {nam_phan_tich}</span></div></div>
                <div class="bg-white p-5 rounded-xl border border-[#c4c5d7] card-shadow"><span class="text-[10px] text-[#434655] uppercase font-bold">Active</span><div class="mt-2"><span class="text-3xl font-bold text-[#1b1c1d]">{active_interns}</span></div></div>
                <div class="bg-white p-5 rounded-xl border border-[#c4c5d7] card-shadow"><span class="text-[10px] text-[#434655] uppercase font-bold">Completed</span><div class="mt-2"><span class="text-3xl font-bold text-[#1b1c1d]">{completed_interns}</span></div></div>
                <div class="bg-white p-5 rounded-xl border border-[#c4c5d7] card-shadow"><span class="text-[10px] text-[#434655] uppercase font-bold">Converted</span><div class="mt-2 flex items-baseline gap-2"><span class="text-3xl font-bold text-[#335bae]">{converted_interns}</span><span class="text-[10px] text-[#335bae] font-bold">{conversion_rate}% Rate</span></div></div>
                <div class="bg-white p-5 rounded-xl border border-[#c4c5d7] card-shadow"><span class="text-[10px] text-[#ba1a1a] uppercase font-bold">Early Leave</span><div class="mt-2"><span class="text-3xl font-bold text-[#ba1a1a]">{early_leave}</span></div></div>
            </div>
        </section>
        <section class="space-y-6 mb-8 animate-fade-in delay-2">
            <h3 class="text-xl font-bold text-[#0045d3] flex items-center gap-2"><span class="material-symbols-outlined">analytics</span> II. Intern Structure</h3>
            <div class="grid grid-cols-1 md:grid-cols-3 gap-6 mb-6">
                <div class="bg-white p-6 rounded-xl border border-[#c4c5d7] card-shadow"><h4 class="font-bold text-sm mb-6 text-[#1b1c1d]">Distribution by Department</h4><div class="space-y-4">{dept_int_html}</div></div>
                <div class="bg-white p-6 rounded-xl border border-[#c4c5d7] card-shadow"><h4 class="font-bold text-sm mb-6 text-[#1b1c1d]">Distribution by Team</h4><div class="grid grid-cols-2 grid-rows-2 gap-2 h-40">{team_int_html}</div></div>
                <div class="bg-white p-6 rounded-xl border border-[#c4c5d7] card-shadow"><h4 class="font-bold text-sm mb-6 text-[#1b1c1d]">Distribution by Major</h4>{major_html}</div>
            </div>
            <div class="grid grid-cols-1 md:grid-cols-3 gap-6">
                <div class="bg-white p-6 rounded-xl border border-[#c4c5d7] card-shadow"><h4 class="font-bold text-sm mb-6 text-[#1b1c1d]">Top Universities</h4><div class="space-y-2">{uni_int_html}</div></div>
                <div class="bg-white p-6 rounded-xl border border-[#c4c5d7] card-shadow flex flex-col justify-center items-center"><h4 class="font-bold text-sm mb-6 text-[#1b1c1d] w-full text-left">Gender Distribution</h4>
                    <div class="flex items-center gap-8"><div class="text-center"><div class="w-14 h-14 rounded-full bg-[#0045d3]/10 flex items-center justify-center text-[#0045d3] mb-2"><span class="material-symbols-outlined text-2xl">male</span></div><div class="font-bold text-lg">{p_nam_int}%</div><div class="text-[10px] text-[#747686] uppercase">Male</div></div>
                    <div class="text-center"><div class="w-14 h-14 rounded-full bg-[#335bae]/10 flex items-center justify-center text-[#335bae] mb-2"><span class="material-symbols-outlined text-2xl">female</span></div><div class="font-bold text-lg">{100-p_nam_int if total_interns>0 else 0}%</div><div class="text-[10px] text-[#747686] uppercase">Female</div></div></div>
                </div>
                <div class="bg-white p-6 rounded-xl border border-[#c4c5d7] card-shadow"><h4 class="font-bold text-sm mb-6 text-[#1b1c1d]">Age Distribution</h4><div class="h-32 flex items-end gap-4 justify-center border-b border-[#efedee]">{age_int_html}</div></div>
            </div>
        </section>
        <section class="space-y-6 animate-fade-in delay-3">
            <h3 class="text-xl font-bold text-[#0045d3] flex items-center gap-2"><span class="material-symbols-outlined">assignment_ind</span> III. Insights & Alerts</h3>
            <div class="grid grid-cols-1 lg:grid-cols-12 gap-6">
                <div class="lg:col-span-8 bg-white p-6 rounded-xl border border-[#c4c5d7] card-shadow">
                    <h4 class="font-bold text-sm mb-6 text-[#1b1c1d]">Top Universities by Conversion Quality</h4>
                    <div class="overflow-x-auto"><table class="w-full text-left text-[11px]"><thead class="bg-[#f5f3f4]"><tr><th class="p-3 uppercase text-[#434655]">University</th><th class="p-3 uppercase text-[#434655]">Total Interns</th><th class="p-3 uppercase text-[#434655]">Avg Score</th><th class="p-3 uppercase text-[#434655]">Quality Rank</th></tr></thead><tbody class="divide-y divide-[#e2e8f0]"><tr><td class="p-3 font-bold text-[#1b1c1d]">HUST</td><td class="p-3">45</td><td class="p-3">4.8</td><td class="p-3 text-[#0045d3] font-bold">Elite</td></tr><tr><td class="p-3 font-bold text-[#1b1c1d]">PTIT</td><td class="p-3">18</td><td class="p-3">4.6</td><td class="p-3 text-[#335bae] font-bold">High</td></tr><tr><td class="p-3 font-bold text-[#1b1c1d]">FTU</td><td class="p-3">28</td><td class="p-3">4.3</td><td class="p-3 text-[#747686] font-bold">Medium</td></tr></tbody></table></div>
                </div>
                <div class="lg:col-span-4 bg-[#ffdad6]/20 p-6 rounded-xl border border-[#ba1a1a]/20 card-shadow">
                    <h4 class="font-bold text-sm mb-6 text-[#ba1a1a] flex items-center gap-2"><span class="material-symbols-outlined">report_problem</span> High-Risk Areas</h4>
                    <div class="space-y-4"><div class="p-3 bg-white rounded border border-[#ba1a1a]/20 flex items-start gap-3"><div class="w-8 h-8 rounded-full bg-[#ba1a1a]/10 flex items-center justify-center text-[#ba1a1a] shrink-0"><span class="material-symbols-outlined text-sm">logout</span></div><div><div class="font-bold text-[11px] text-[#ba1a1a]">QA Early Leave Spike</div><p class="text-[10px] text-[#434655] mt-1">Tỷ lệ nghỉ sớm tại phòng QA tăng mạnh.</p></div></div></div>
                </div>
            </div>
        </section>
    </body></html>
    """
    components.html(html_intern, height=1400, scrolling=True)


# ==========================================
# TRANG 4: ATTRITION & RECRUITMENT
# ==========================================
elif page == "Attrition & Recruitment":
    
    html_p3 = f"""
    <!DOCTYPE html><html lang="vi"><head>{TAILWIND_HEAD}</head><body class="p-4 md:p-6"><div class="max-w-[1600px] mx-auto animate-fade-in">
        <div class="mb-6"><h1 class="text-3xl font-bold text-[#1b1c1d]">Báo cáo Nghỉ việc & Tuyển dụng</h1><p class="text-[#434655] mt-1 text-sm">Phân tích chi tiết biến động nhân sự ({nam_phan_tich}).</p></div>
        <div class="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-6 gap-4 mb-2 delay-1 animate-fade-in">
            <div class="bg-white p-5 rounded-xl border border-[#c4c5d7] flex flex-col justify-between shadow-sm hover-card"><span class="text-[11px] font-bold text-[#434655] uppercase tracking-wider">Tuyển mới</span><div class="mt-2"><span class="text-3xl font-bold text-[#1b1c1d]">{hires_count}</span></div><p class="text-[11px] text-[#434655] mt-2 border-t border-[#f5f3f4] pt-2">Trong năm {nam_phan_tich}</p></div>
            <div class="bg-white p-5 rounded-xl border border-[#c4c5d7] flex flex-col justify-between shadow-sm hover-card"><span class="text-[11px] font-bold text-[#434655] uppercase tracking-wider">Nghỉ việc</span><div class="mt-2"><span class="text-3xl font-bold text-[#ba1a1a]">{terms_count}</span></div><p class="text-[11px] text-[#434655] mt-2 border-t border-[#f5f3f4] pt-2">Trong năm {nam_phan_tich}</p></div>
            <div class="bg-white p-5 rounded-xl border border-[#c4c5d7] flex flex-col justify-between shadow-sm hover-card"><span class="text-[11px] font-bold text-[#434655] uppercase tracking-wider">Turnover Rate</span><div class="mt-2"><span class="text-3xl font-bold text-[#ba1a1a]">{turnover_rate}%</span></div><p class="text-[11px] text-[#434655] mt-2 border-t border-[#f5f3f4] pt-2">Tỷ lệ nghỉ việc (HĐ chính thức)</p></div>
            <div class="bg-white p-5 rounded-xl border border-[#c4c5d7] flex flex-col justify-between shadow-sm hover-card"><span class="text-[11px] font-bold text-[#434655] uppercase tracking-wider">Hiring Rate</span><div class="mt-2"><span class="text-3xl font-bold text-[#3260ec]">{hiring_rate}%</span></div><p class="text-[11px] text-[#434655] mt-2 border-t border-[#f5f3f4] pt-2">Tỷ lệ tuyển dụng</p></div>
            <div class="bg-white p-5 rounded-xl border border-[#c4c5d7] flex flex-col justify-between shadow-sm hover-card"><span class="text-[11px] font-bold text-[#434655] uppercase tracking-wider">Net Change</span><div class="mt-2"><span class="text-3xl font-bold {'text-[#3260ec]' if net_change>=0 else 'text-[#ba1a1a]'}">{'+' if net_change>0 else ''}{net_change}</span></div><p class="text-[11px] text-[#434655] mt-2 border-t border-[#f5f3f4] pt-2">Tăng/giảm thuần</p></div>
            <div class="bg-white p-5 rounded-xl border border-[#c4c5d7] flex flex-col justify-between shadow-sm hover-card"><span class="text-[11px] font-bold text-[#434655] uppercase tracking-wider">Retention</span><div class="mt-2"><span class="text-3xl font-bold text-[#1b1c1d]">{retention_rate}%</span></div><p class="text-[11px] text-[#434655] mt-2 border-t border-[#f5f3f4] pt-2">Tỷ lệ giữ chân</p></div>
        </div>
    </div></body></html>
    """
    components.html(html_p3, height=230, scrolling=False)

    st.markdown(create_card("Phân tích Tuyển dụng", "person_add", "#3260ec"), unsafe_allow_html=True)
    r1c1, r1c2 = st.columns([7, 5])
    with r1c1:
        if not df_hires.empty:
            hire_trend = df_hires['Tháng vào làm'].value_counts().sort_index().reset_index()
            hire_trend.columns = ['Tháng', 'Tuyển mới']
            fig_h1 = px.line(hire_trend, x='Tháng', y='Tuyển mới', markers=True, color_discrete_sequence=['#3260ec'])
            fig_h1.update_traces(line=dict(width=3, shape='spline'), marker=dict(size=8), fill='tozeroy', fillcolor='rgba(50, 96, 236, 0.15)')
            chart_wrapper(f"Nhân sự tuyển mới theo tháng ({nam_phan_tich})", fig_h1, height=220)
        
    with r1c2:
        s_nam_vao = pd.to_numeric(df_base['Năm vào làm'], errors='coerce').dropna().astype(int)
        if not s_nam_vao.empty:
            yearly_hires = s_nam_vao.value_counts().sort_index().reset_index()
            yearly_hires.columns = ['Năm', 'Tuyển mới']
            colors = ['#3260ec' if year == nam_phan_tich else '#c4c5d7' for year in yearly_hires['Năm']]
            fig_h_year = px.bar(yearly_hires, x='Năm', y='Tuyển mới', text_auto=True)
            fig_h_year.update_traces(marker_color=colors)
            fig_h_year.update_xaxes(type='category') 
            chart_wrapper("Xu hướng tuyển mới theo Năm", fig_h_year, height=220)

    st.markdown("<br>", unsafe_allow_html=True)
    r2c1, r2c2, r2c3 = st.columns(3)
    with r2c1:
        if not df_hires.empty and 'Phòng ban' in df_hires.columns:
            dept_hires = df_hires[df_hires['Phòng ban'] != 'Chưa cập nhật']['Phòng ban'].value_counts().head(5).reset_index()
            dept_hires.columns = ['Phòng ban', 'Số lượng']
            fig_h2 = px.bar(dept_hires, x='Số lượng', y='Phòng ban', orientation='h', text_auto=True, color_discrete_sequence=['#3260ec'])
            fig_h2.update_yaxes(categoryorder='total ascending')
            chart_wrapper(f"Tuyển mới theo Phòng ban ({nam_phan_tich})", fig_h2, height=220)

    with r2c2:
        if not df_hires.empty and 'Cấp bậc' in df_hires.columns:
            rank_hires = df_hires[df_hires['Cấp bậc']!='Chưa cập nhật']['Cấp bậc'].value_counts().reset_index()
            rank_hires.columns = ['Cấp bậc', 'Số lượng']
            fig_h3 = px.bar(rank_hires, x='Cấp bậc', y='Số lượng', text_auto=True, color_discrete_sequence=['#82a6fe'])
            chart_wrapper(f"Tuyển mới theo Cấp bậc ({nam_phan_tich})", fig_h3, height=220)
        
    with r2c3:
        if not df_hires.empty and 'Công nghệ' in df_hires.columns:
            tech_hires = df_hires[df_hires['Công nghệ']!='Chưa cập nhật']['Công nghệ'].value_counts().reset_index()
            tech_hires.columns = ['Công nghệ', 'Số lượng']
            fig_h4 = px.treemap(tech_hires, path=[px.Constant("Tech"), 'Công nghệ'], values='Số lượng', color='Số lượng', color_continuous_scale='Blues')
            fig_h4.update_layout(margin=dict(t=0, l=0, r=0, b=0))
            chart_wrapper(f"Tuyển mới theo Công nghệ ({nam_phan_tich})", fig_h4, height=220)

    # --- PHÂN TÍCH NGHỈ VIỆC ---
    st.markdown("<br><hr style='border-color:#e2e8f0;'><br>", unsafe_allow_html=True)
    st.markdown(create_card("Phân tích Nghỉ việc", "trending_down", "#ba1a1a"), unsafe_allow_html=True)
    
    r3c1, r3c2, r3c3 = st.columns([4, 5, 3])
    with r3c1:
        if not df_terms.empty:
            term_trend = df_terms['Tháng nghỉ'].value_counts().sort_index().reset_index()
            term_trend.columns = ['Tháng', 'Nghỉ việc']
            fig_t1 = px.line(term_trend, x='Tháng', y='Nghỉ việc', markers=True, color_discrete_sequence=['#ba1a1a'])
            fig_t1.update_traces(line=dict(width=3, shape='spline'), marker=dict(size=8), fill='tozeroy', fillcolor='rgba(186, 26, 26, 0.1)')
            chart_wrapper(f"Nghỉ việc theo tháng ({nam_phan_tich})", fig_t1, height=250)
        
    with r3c2:
        if not df_terms.empty and 'Lý do nghỉ việc' in df_terms.columns:
            reason_terms = df_terms[df_terms['Lý do nghỉ việc']!='Chưa cập nhật']['Lý do nghỉ việc'].value_counts().reset_index()
            reason_terms.columns = ['Lý do', 'Số lượng']
            reason_terms['Lý do (Ngắn)'] = reason_terms['Lý do'].apply(lambda x: str(x)[:30] + '...' if len(str(x)) > 30 else str(x))
            fig_t2 = px.pie(reason_terms, names='Lý do (Ngắn)', values='Số lượng', custom_data=['Lý do'], hole=0.5, color_discrete_sequence=['#174597', '#3260ec', '#82a6fe', '#dce1ff', '#c4c5d7'])
            fig_t2.update_traces(textposition='inside', textinfo='percent', hovertemplate="<b>%{customdata[0]}</b><br>Số lượng: %{value}<extra></extra>")
            fig_t2.update_layout(showlegend=True, margin=dict(t=10, b=10, l=0, r=0), legend=dict(title="", orientation="v", yanchor="middle", y=0.5, xanchor="left", x=1.0, font=dict(size=11, color="#434655"), itemwidth=30))
            chart_wrapper(f"Lý do nghỉ việc ({nam_phan_tich})", fig_t2, height=250)
        
    with r3c3:
        if not df_terms.empty and 'Nhóm thâm niên' in df_terms.columns:
            tn_terms = df_terms[df_terms['Nhóm thâm niên']!='Chưa cập nhật']['Nhóm thâm niên'].value_counts().sort_index().reset_index()
            tn_terms.columns = ['Thâm niên', 'Số lượng']
            fig_t3 = px.bar(tn_terms, x='Thâm niên', y='Số lượng', text_auto=True, color_discrete_sequence=['#ffdad6'])
            fig_t3.update_traces(marker_line_color='#ba1a1a', marker_line_width=1.5)
            chart_wrapper(f"Thâm niên người nghỉ ({nam_phan_tich})", fig_t3, height=250)

    # 4 BIỂU ĐỒ NGHỈ VIỆC BỔ SUNG 
    r4c1, r4c2, r4c3, r4c4 = st.columns(4)
    with r4c1:
        if not df_terms.empty and 'Phòng ban' in df_terms.columns:
            dept_terms = df_terms[df_terms['Phòng ban'] != 'Chưa cập nhật']['Phòng ban'].value_counts().head(5).reset_index()
            dept_terms.columns = ['Phòng ban', 'Số lượng']
            fig_t4 = px.bar(dept_terms, x='Số lượng', y='Phòng ban', orientation='h', text_auto=True, color_discrete_sequence=['#ffdad6'])
            fig_t4.update_traces(marker_line_color='#ba1a1a', marker_line_width=1.5)
            fig_t4.update_yaxes(categoryorder='total ascending') 
            chart_wrapper(f"Nghỉ việc theo Phòng ban", fig_t4, height=250)
            
    with r4c2:
        if not df_terms.empty and 'Cấp bậc' in df_terms.columns:
            rank_terms = df_terms[df_terms['Cấp bậc']!='Chưa cập nhật']['Cấp bậc'].value_counts().reset_index()
            rank_terms.columns = ['Cấp bậc', 'Số lượng']
            fig_t5 = px.bar(rank_terms, x='Cấp bậc', y='Số lượng', text_auto=True, color_discrete_sequence=['#ffb596'])
            fig_t5.update_traces(marker_line_color='#ba4800', marker_line_width=1.5)
            chart_wrapper(f"Nghỉ việc theo Cấp bậc", fig_t5, height=250)
            
    with r4c3:
        if not df_terms.empty and 'Nhóm tuổi' in df_terms.columns:
            age_terms = df_terms[df_terms['Nhóm tuổi']!='Chưa cập nhật']['Nhóm tuổi'].value_counts().sort_index().reset_index()
            age_terms.columns = ['Nhóm tuổi', 'Số lượng']
            fig_t6 = px.bar(age_terms, x='Nhóm tuổi', y='Số lượng', text_auto=True, color_discrete_sequence=['#c4c5d7'])
            fig_t6.update_traces(marker_line_color='#434655', marker_line_width=1.5)
            chart_wrapper(f"Độ tuổi Nghỉ việc", fig_t6, height=250)
            
    with r4c4:
        if not df_terms.empty and 'Công nghệ' in df_terms.columns:
            tech_terms = df_terms[df_terms['Công nghệ']!='Chưa cập nhật']['Công nghệ'].value_counts().head(5).reset_index()
            tech_terms.columns = ['Công nghệ', 'Số lượng']
            fig_t7 = px.bar(tech_terms, x='Công nghệ', y='Số lượng', text_auto=True, color_discrete_sequence=['#ba1a1a'])
            chart_wrapper(f"Công nghệ thiếu ổn định", fig_t7, height=250)

    # --- SO SÁNH & BIẾN ĐỘNG ---
    st.markdown("<br><hr style='border-color:#e2e8f0;'><br>", unsafe_allow_html=True)
    r5c1, r5c2 = st.columns([6, 6])
    
    if not df_hires.empty or not df_terms.empty:
        h_tr = df_hires['Tháng vào làm'].value_counts().rename('Tuyển mới') if not df_hires.empty else pd.Series(dtype=int)
        t_tr = df_terms['Tháng nghỉ'].value_counts().rename('Nghỉ việc') if not df_terms.empty else pd.Series(dtype=int)
        trend_df = pd.merge(h_tr, t_tr, left_index=True, right_index=True, how='outer').fillna(0).sort_index().reset_index()
        trend_df.columns = ['Tháng', 'Tuyển mới', 'Nghỉ việc']
        trend_df['Tháng'] = "Tháng " + trend_df['Tháng'].astype(int).astype(str)
        trend_df['Net'] = trend_df['Tuyển mới'] - trend_df['Nghỉ việc']
        
        with r5c1:
            fig_vs = go.Figure()
            fig_vs.add_trace(go.Bar(x=trend_df['Tháng'], y=trend_df['Tuyển mới'], name='Tuyển mới', marker_color='#3260ec', width=0.35))
            fig_vs.add_trace(go.Bar(x=trend_df['Tháng'], y=trend_df['Nghỉ việc'], name='Nghỉ việc', marker_color='#ba1a1a', width=0.35))
            fig_vs.update_layout(barmode='group', legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
            chart_wrapper(f"Tuyển mới vs Nghỉ việc theo tháng ({nam_phan_tich})", fig_vs, height=280)
            
        with r5c2:
            fig_wf = go.Figure(go.Waterfall(x=trend_df['Tháng'], y=trend_df['Net'], text=trend_df['Net'], textposition="outside", measure=["relative"] * len(trend_df), decreasing={"marker":{"color":"#ba1a1a"}}, increasing={"marker":{"color":"#3260ec"}}))
            fig_wf.update_layout(waterfallgap=0.3)
            chart_wrapper(f"Biến động nhân sự thuần ({nam_phan_tich})", fig_wf, height=280)

    st.markdown("<div style='background:white; padding:24px; border-radius:12px; border:1px solid #c4c5d7;' class='card-shadow'><div style='font-family:Inter; font-size:14px; font-weight:700; color:#1b1c1d; margin-bottom:15px;'>Danh sách chi tiết nhân sự nghỉ việc</div>", unsafe_allow_html=True)
    cols_to_show = ['MSNV', 'Họ tên', 'Phòng ban', 'Chức danh', 'Cấp bậc', 'Ngày vào làm', 'Ngày làm việc cuối cùng', 'Lý do nghỉ việc']
    st.dataframe(df_terms[[c for c in cols_to_show if c in df_terms.columns]], use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)

# ==========================================
# TRANG 5: CONTRACT & HR ALERTS
# ==========================================
elif page == "Contract & HR Alerts":

    today = datetime.datetime.now()
    current_month = today.month
    next_30 = today + datetime.timedelta(days=30)
    next_60 = today + datetime.timedelta(days=60)
    next_90 = today + datetime.timedelta(days=90)

    st.markdown("<h1 style='color:#1b1c1d; font-family:Inter; font-size:30px; font-weight:700; margin-bottom:5px;'>Hợp đồng & Cảnh báo</h1><p style='color:#434655; font-family:Inter; font-size:14px; margin-bottom:20px;'>Quản lý hiệu lực hợp đồng và theo dõi các thông báo nhân sự quan trọng.</p>", unsafe_allow_html=True)
    
    col_filter1, col_filter2 = st.columns(2)
    with col_filter1:
        alert_filter = st.selectbox("⏳ Khoảng thời gian cảnh báo hợp đồng:", ["30 ngày tới", "60 ngày tới", "90 ngày tới"])
        if alert_filter == "30 ngày tới": target_date = next_30
        elif alert_filter == "60 ngày tới": target_date = next_60
        else: target_date = next_90
        
    with col_filter2:
        month_options = [f"Tháng {i}" for i in range(1, 13)]
        selected_month_name = st.selectbox("🎂 Chọn tháng sinh nhật:", month_options, index=current_month - 1)
        selected_month_num = int(selected_month_name.replace("Tháng ", ""))

    df_current_active = df_base[df_base['Tình trạng'] != 'OFF'].copy()
    current_total_emp = len(df_current_active)

    def phan_loai_hd_trang4(row):
        hd_text = str(row.get('HĐ hiện tại', '')).lower()
        if 'không xác định' in hd_text or 'vô thời hạn' in hd_text: return 'Vô thời hạn'
        elif any(kw in hd_text for kw in ['thử việc', 'ctv', 'cộng tác', 'thực tập', 'intern']): return 'Thử việc / CTV'
        else: return 'Xác định thời hạn'

    if not df_current_active.empty:
        df_current_active['Loại HĐ T4'] = df_current_active.apply(phan_loai_hd_trang4, axis=1)
        hd_vth = len(df_current_active[df_current_active['Loại HĐ T4'] == 'Vô thời hạn'])
        thu_viec_ctv = len(df_current_active[df_current_active['Loại HĐ T4'] == 'Thử việc / CTV'])
        hd_xdth = len(df_current_active[df_current_active['Loại HĐ T4'] == 'Xác định thời hạn'])
    else: hd_vth = thu_viec_ctv = hd_xdth = 0

    pct_vth = int(round((hd_vth / current_total_emp) * 100, 0)) if current_total_emp > 0 else 0
    pct_tv = int(round((thu_viec_ctv / current_total_emp) * 100, 0)) if current_total_emp > 0 else 0
    pct_xdth = 100 - pct_vth - pct_tv if current_total_emp > 0 else 0 

    if 'HĐ hiện tại đến' in df_current_active.columns:
        ngay_het_han = pd.to_datetime(df_current_active['HĐ hiện tại đến'], errors='coerce')
        df_expiring = df_current_active[(ngay_het_han >= today) & (ngay_het_han <= target_date)].copy()
        df_expiring['HĐ hiện tại đến'] = pd.to_datetime(df_expiring['HĐ hiện tại đến'], errors='coerce')
        df_expiring = df_expiring.sort_values(by='HĐ hiện tại đến')
    else: df_expiring = pd.DataFrame()

    exp_html = ""
    for _, row in df_expiring.head(50).iterrows():
        name = row.get('Họ tên', row.get('MSNV', 'Unknown'))
        name_str = str(name) if pd.notna(name) else "Unknown"
        initial = name_str[0].upper() if name_str else ""
        msnv = row.get('MSNV', '')
        hd_type = row.get('HĐ hiện tại', 'Xác định thời hạn')
        end_date = row['HĐ hiện tại đến'].strftime('%d/%m/%Y')
        days_left = (row['HĐ hiện tại đến'] - today).days
        exp_html += f'<tr class="hover:bg-[#f5f3f4] transition-colors border-b border-[#e2e8f0]"><td class="px-6 py-4 whitespace-nowrap"><div class="flex items-center gap-3"><div class="w-8 h-8 rounded-full bg-[#ffdad6] text-[#ba1a1a] flex items-center justify-center font-bold text-xs">{initial}</div><div><p class="text-sm font-bold text-[#1b1c1d]">{name_str}</p><p class="text-[10px] text-[#747686]">ID: {msnv}</p></div></div></td><td class="px-6 py-4 text-sm text-[#434655] whitespace-nowrap">{hd_type}</td><td class="px-6 py-4 text-sm text-[#434655] whitespace-nowrap">{end_date}</td><td class="px-6 py-4 whitespace-nowrap"><span class="inline-flex items-center px-2 py-1 rounded bg-[#ffdad6] text-[#93000a] text-[10px] font-bold"><span class="material-symbols-outlined text-[14px] mr-1">timer</span> {days_left} ngày</span></td><td class="px-6 py-4 text-right whitespace-nowrap"><button class="bg-[#0045d3] text-white text-xs px-4 py-1.5 rounded-lg hover:bg-opacity-90 transition">Gia hạn</button></td></tr>'
    
    if df_expiring.empty: exp_html = "<tr><td colspan='5' class='text-center py-10 text-sm text-[#747686]'>Không có hợp đồng sắp hết hạn.</td></tr>"

    df_current_active['Tháng sinh'] = pd.to_datetime(df_current_active['Ngày sinh'], errors='coerce').dt.month
    df_bdays = df_current_active[df_current_active['Tháng sinh'] == selected_month_num].copy()
    
    bd_html = ""
    for _, row in df_bdays.head(50).iterrows():
        name = row.get('Họ tên', row.get('MSNV', 'Unknown'))
        name_str = str(name) if pd.notna(name) else "Unknown"
        initial = name_str[0].upper() if name_str else ""
        dept = row.get('Phòng ban', '')
        day = pd.to_datetime(row['Ngày sinh']).day if pd.notna(row['Ngày sinh']) else ""
        bd_html += f'<div class="flex items-center justify-between p-3 hover:bg-[#f5f3f4] rounded-lg transition-colors cursor-pointer mb-1 border border-transparent hover:border-[#e2e8f0]"><div class="flex items-center gap-3"><div class="w-10 h-10 rounded-full bg-[#dce1ff] text-[#0045d3] flex items-center justify-center font-bold text-lg">{initial}</div><div><p class="text-sm font-bold text-[#1b1c1d]">{name_str}</p><p class="text-[10px] text-[#747686]">{day} {selected_month_name} • {dept}</p></div></div><span class="material-symbols-outlined text-[#3260ec] text-[20px]">cake</span></div>'
    
    if df_bdays.empty: bd_html = f"<div class='text-center py-10 text-sm text-[#747686]'>Không có sinh nhật trong {selected_month_name}.</div>"

    r1c1, r1c2 = st.columns([35, 65]) 
    with r1c1:
        html_donut = f"""<!DOCTYPE html><html lang="vi"><head>{TAILWIND_HEAD}<style>.donut-chart {{ background: conic-gradient(#3260ec 0% {pct_xdth}%, #82a6fe {pct_xdth}% {pct_xdth+pct_vth}%, #dce1ff {pct_xdth+pct_vth}% 100%); border-radius: 50%; }}</style></head><body style="background:transparent; margin:0; padding:4px;"><div class="bg-white border border-[#c4c5d7] rounded-xl p-6 flex flex-col h-[400px] card-shadow animate-fade-in delay-1"><div class="flex justify-between items-center mb-6"><div class="text-base font-bold text-[#1b1c1d]">Phân loại hợp đồng</div></div><div class="flex-1 flex flex-col items-center justify-center"><div class="relative w-44 h-44 mb-6"><div class="donut-chart w-full h-full"></div><div class="absolute inset-4 bg-white rounded-full flex flex-col items-center justify-center"><span class="text-4xl font-bold text-[#1b1c1d]">{current_total_emp}</span><span class="text-[11px] text-[#747686] uppercase tracking-wider">Tổng số</span></div></div><div class="w-full space-y-2"><div class="flex justify-between items-center"><div class="flex items-center gap-2"><div class="w-3 h-3 rounded-full bg-[#82a6fe]"></div><span class="text-sm text-[#434655]">Vô thời hạn</span></div><span class="font-bold text-[#1b1c1d]">{pct_vth}%</span></div><div class="flex justify-between items-center"><div class="flex items-center gap-2"><div class="w-3 h-3 rounded-full bg-[#3260ec]"></div><span class="text-sm text-[#434655]">Xác định thời hạn</span></div><span class="font-bold text-[#1b1c1d]">{pct_xdth}%</span></div><div class="flex justify-between items-center"><div class="flex items-center gap-2"><div class="w-3 h-3 rounded-full bg-[#dce1ff]"></div><span class="text-sm text-[#434655]">Thử việc / CTV</span></div><span class="font-bold text-[#1b1c1d]">{pct_tv}%</span></div></div></div></div></body></html>"""
        components.html(html_donut, height=420)

    with r1c2:
        html_table = f"""<!DOCTYPE html><html lang="vi"><head>{TAILWIND_HEAD}</head><body style="background:transparent; margin:0; padding:4px;"><div class="bg-white border border-[#c4c5d7] rounded-xl flex flex-col card-shadow h-[400px] animate-fade-in delay-1"><div class="p-6 border-b border-[#e2e8f0] flex justify-between items-center shrink-0"><div class="text-base font-bold text-[#1b1c1d]">Hợp đồng sắp hết hạn ({len(df_expiring)} NV)</div></div><div class="flex-1 overflow-auto custom-scrollbar"><table class="w-full text-left"><thead class="bg-[#f5f3f4] sticky top-0 z-10 border-b border-[#e2e8f0] shadow-sm"><tr><th class="px-6 py-3 text-[11px] text-[#434655] uppercase tracking-wider font-bold">Nhân viên</th><th class="px-6 py-3 text-[11px] text-[#434655] uppercase tracking-wider font-bold">Loại hợp đồng</th><th class="px-6 py-3 text-[11px] text-[#434655] uppercase tracking-wider font-bold">Ngày hết hạn</th><th class="px-6 py-3 text-[11px] text-[#434655] uppercase tracking-wider font-bold">Còn lại</th><th class="px-6 py-3 text-[11px] text-[#434655] uppercase tracking-wider text-right font-bold">Hành động</th></tr></thead><tbody>{exp_html}</tbody></table></div></div></body></html>"""
        components.html(html_table, height=420)

    r2c1, r2c2 = st.columns([65, 35])
    with r2c1:
        html_alerts = f"""<!DOCTYPE html><html lang="vi"><head>{TAILWIND_HEAD}</head><body style="background:transparent; margin:0; padding:4px;"><div class="bg-[#f8fafc] border border-dashed border-[#c4c5d7] rounded-xl p-6 card-shadow flex flex-col items-center justify-center text-center h-[380px] animate-fade-in delay-2"><div class="w-16 h-16 bg-[#dce1ff] rounded-full flex items-center justify-center mb-4"><span class="material-symbols-outlined text-[#3260ec] text-3xl">build</span></div><div class="text-lg font-bold text-[#1b1c1d] mb-2">Cảnh báo nhân sự (Compliance & Alerts)</div><p class="text-sm text-[#747686] max-w-md">Khu vực này đang được chuẩn bị để đồng bộ với dữ liệu Hồ sơ pháp lý trong giai đoạn tới.</p></div></body></html>"""
        components.html(html_alerts, height=400)

    with r2c2:
        html_bdays = f"""<!DOCTYPE html><html lang="vi"><head>{TAILWIND_HEAD}</head><body style="background:transparent; margin:0; padding:4px;"><div class="bg-white border border-[#c4c5d7] rounded-xl p-6 card-shadow flex flex-col h-[380px] animate-fade-in delay-2"><div class="flex justify-between items-center mb-4 shrink-0"><div class="text-base font-bold text-[#1b1c1d]">Sinh nhật {selected_month_name.lower()}</div><span class="text-xs text-[#0045d3] font-bold bg-[#dce1ff] px-2 py-1 rounded-full">{len(df_bdays)} nhân viên</span></div><div class="flex-1 overflow-auto custom-scrollbar space-y-1 pr-1">{bd_html}</div><button class="w-full mt-4 py-2.5 border border-[#c4c5d7] text-sm font-bold rounded-lg hover:bg-[#f5f3f4] transition-colors text-[#1b1c1d] shrink-0">Gửi lời chúc hàng loạt</button></div></body></html>"""
        components.html(html_bdays, height=400)
