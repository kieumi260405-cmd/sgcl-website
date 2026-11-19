
import streamlit as st

st.set_page_config(
    page_title="SGCL | Ship Agency Management",
    page_icon="🚢",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
:root {
    --navy: #0B1F3A;
    --navy-light: #15325C;
    --blue: #2563EB;
    --blue-soft: #3B82F6;
    --bg: #F1F5F9;
    --white: #FFFFFF;
    --text: #1E293B;
    --muted: #64748B;
    --border: #E2E8F0;
    --card-shadow: 0 4px 6px -1px rgba(15, 23, 42, 0.06), 0 2px 4px -2px rgba(15, 23, 42, 0.04);
    --card-shadow-hover: 0 10px 15px -3px rgba(15, 23, 42, 0.08), 0 4px 6px -4px rgba(15, 23, 42, 0.05);
}

.stApp {
    background-color: var(--bg);
    color: var(--text);
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
}

header[data-testid="stHeader"] {
    background: transparent;
}

section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #0B1F3A 0%, #0F2A4A 100%);
    border-right: 1px solid rgba(255,255,255,0.06);
}

section[data-testid="stSidebar"] > div {
    background: transparent;
    padding-top: 1.2rem;
}

section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"],
section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] *,
section[data-testid="stSidebar"] label,
section[data-testid="stSidebar"] p,
section[data-testid="stSidebar"] span,
section[data-testid="stSidebar"] .stRadio label {
    color: #E2E8F0 !important;
}

section[data-testid="stSidebar"] .stRadio > div {
    gap: 0.3rem;
}

section[data-testid="stSidebar"] .stRadio label {
    background: transparent;
    padding: 0.6rem 0.9rem !important;
    border-radius: 8px;
    transition: all 0.2s ease;
    font-weight: 500;
    font-size: 0.95rem;
}

section[data-testid="stSidebar"] .stRadio label:hover {
    background: rgba(255,255,255,0.08) !important;
}

.logo-box {
    padding: 0.3rem 0.5rem 1rem 0.5rem;
}

.logo-title {
    color: white;
    font-size: 28px;
    font-weight: 800;
    letter-spacing: 0.5px;
    margin-bottom: 2px;
}

.logo-subtitle {
    color: #94A3B8;
    font-size: 11px;
    font-weight: 600;
    letter-spacing: 1.8px;
    text-transform: uppercase;
}

.top-header {
    background: var(--white);
    padding: 1.35rem 1.7rem;
    border: 1px solid var(--border);
    border-radius: 14px;
    margin-bottom: 1.6rem;
    box-shadow: var(--card-shadow);
}

.top-label {
    color: var(--muted);
    font-size: 12px;
    font-weight: 600;
    letter-spacing: 0.5px;
    text-transform: uppercase;
}

.top-title {
    color: var(--navy);
    font-size: 1.35rem;
    font-weight: 800;
    letter-spacing: -0.3px;
    margin-top: 3px;
}

.top-desc {
    color: var(--muted);
    font-size: 0.95rem;
    margin-top: 3px;
}

.page-title {
    color: var(--navy);
    font-size: 1.65rem;
    font-weight: 800;
    letter-spacing: -0.4px;
    margin-bottom: 0.3rem;
}

.page-subtitle {
    color: var(--muted);
    font-size: 1.05rem;
    font-weight: 500;
    line-height: 1.5;
    margin-bottom: 1.4rem;
}

div[data-testid="stMetric"] {
    background: var(--white);
    padding: 1.35rem 1.4rem;
    border: 1px solid var(--border);
    border-radius: 14px;
    box-shadow: var(--card-shadow);
    transition: all 0.25s ease;
}

div[data-testid="stMetric"]:hover {
    box-shadow: var(--card-shadow-hover);
    transform: translateY(-2px);
}

div[data-testid="stMetric"] label {
    color: var(--muted) !important;
    font-size: 0.82rem !important;
    font-weight: 600 !important;
    text-transform: uppercase;
    letter-spacing: 0.4px;
}

div[data-testid="stMetric"] [data-testid="stMetricValue"] {
    color: var(--navy) !important;
    font-size: 1.9rem !important;
    font-weight: 800 !important;
}

.stApp h1, .stApp h2, .stApp h3 {
    color: var(--navy) !important;
    font-weight: 700 !important;
}

.stButton > button {
    background: linear-gradient(135deg, var(--navy) 0%, var(--navy-light) 100%);
    color: white !important;
    border: none;
    border-radius: 9px;
    font-weight: 600;
    min-height: 42px;
    padding: 0.5rem 1.25rem;
    transition: all 0.2s ease;
    box-shadow: 0 2px 4px rgba(11, 31, 58, 0.2);
}

.stButton > button:hover {
    background: linear-gradient(135deg, var(--blue) 0%, var(--blue-soft) 100%);
    color: white !important;
    box-shadow: 0 4px 12px rgba(37, 99, 235, 0.35);
    transform: translateY(-1px);
}

div[data-testid="stAlert"] {
    border-radius: 12px;
    border: none;
    box-shadow: var(--card-shadow);
}

.process-card {
    background: var(--white);
    border: 1px solid var(--border);
    border-radius: 14px;
    padding: 1.35rem 1.1rem;
    text-align: center;
    height: 100%;
    box-shadow: var(--card-shadow);
    transition: all 0.25s ease;
}

.process-card:hover {
    box-shadow: var(--card-shadow-hover);
    transform: translateY(-3px);
    border-color: #BFDBFE;
}

.process-number {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 36px;
    height: 36px;
    background: linear-gradient(135deg, var(--navy) 0%, var(--blue) 100%);
    color: white;
    font-weight: 800;
    font-size: 0.95rem;
    border-radius: 10px;
    margin-bottom: 0.8rem;
}

.process-title {
    color: var(--navy);
    font-weight: 700;
    font-size: 1.05rem;
    margin-bottom: 0.35rem;
}

.process-desc {
    color: var(--muted);
    font-size: 0.9rem;
    line-height: 1.45;
}

.content-card {
    background: var(--white);
    border: 1px solid var(--border);
    border-radius: 14px;
    padding: 1.4rem 1.6rem;
    box-shadow: var(--card-shadow);
    margin-bottom: 1.2rem;
}

section[data-testid="stFileUploader"] {
    background: var(--white);
    border: 2px dashed #CBD5E1;
    border-radius: 14px;
    padding: 0.8rem;
}

section[data-testid="stFileUploader"]:hover {
    border-color: var(--blue);
    background: #F8FAFC;
}

.footer {
    text-align: center;
    color: #94A3B8;
    font-size: 0.8rem;
    padding: 0.8rem 0 0.3rem 0;
}

hr {
    border: none;
    border-top: 1px solid var(--border);
    margin: 1.6rem 0;
}
</style>
""", unsafe_allow_html=True)

# ===================== SIDEBAR =====================
with st.sidebar:
    st.markdown("""
    <div class="logo-box">
        <div class="logo-title">🚢 SGCL</div>
        <div class="logo-subtitle">Ship Agency Management</div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("##### 📋 DANH MỤC CHỨC NĂNG")

    menu = st.radio(
        "Điều hướng",
        [
            "🏠 Tổng quan",
            "🚢 Quản lý hồ sơ tàu",
            "📤 Tải lên chứng từ",
            "🔍 Trích xuất dữ liệu",
            "✅ Kiểm tra dữ liệu",
            "📊 Xuất Excel NSW"
        ],
        label_visibility="collapsed"
    )

    st.markdown("---")
    st.caption("HỆ THỐNG SGCL")
    st.caption("Phiên bản thử nghiệm 1.0")

# ===================== HEADER =====================
st.markdown("""
<div class="top-header">
    <div class="top-label">SGCL · Ship Agency Management</div>
    <div class="top-title">Hệ thống tự động hóa nhập liệu chứng từ</div>
    <div class="top-desc">Phục vụ nghiệp vụ khai báo tàu nhập cảnh · Tự động OCR · Xuất NSW</div>
</div>
""", unsafe_allow_html=True)

def page_heading(title, description):
    st.markdown(f'<div class="page-title">{title}</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="page-subtitle">{description}</div>', unsafe_allow_html=True)

# ===================== TỔNG QUAN =====================
if menu == "🏠 Tổng quan":
    page_heading("Tổng quan hệ thống", "Theo dõi tình hình xử lý hồ sơ tàu và tiến độ nghiệp vụ")

    col1, col2, col3 = st.columns(3)
    col1.metric("Tổng hồ sơ tàu", "0")
    col2.metric("Hồ sơ đang xử lý", "0")
    col3.metric("Hồ sơ đã hoàn thành", "0")

    st.markdown("")

    st.markdown("""
    <div class="content-card">
        <h3 style="margin-top:0; color:#0B1F3A;">Chào mừng đến với hệ thống SGCL</h3>
        <p style="color:#475569; line-height:1.6; margin-bottom:0;">
            Hệ thống hỗ trợ quản lý hồ sơ tàu, tiếp nhận chứng từ, tự động trích xuất và kiểm tra dữ liệu, 
            phục vụ quá trình khai báo tàu nhập cảnh nhanh chóng, chính xác và đồng bộ với biểu mẫu NSW.
        </p>
    </div>
    """, unsafe_allow_html=True)

    st.subheader("Quy trình xử lý hồ sơ")

    c1, c2, c3, c4 = st.columns(4)

    with c1:
        st.markdown("""
        <div class="process-card">
            <div class="process-number">01</div>
            <div class="process-title">Tiếp nhận</div>
            <div class="process-desc">Tải chứng từ lên hệ thống (PDF, ảnh, Excel...)</div>
        </div>
        """, unsafe_allow_html=True)

    with c2:
        st.markdown("""
        <div class="process-card">
            <div class="process-number">02</div>
            <div class="process-title">Trích xuất</div>
            <div class="process-desc">Nhận diện thông tin tự động bằng OCR</div>
        </div>
        """, unsafe_allow_html=True)

    with c3:
        st.markdown("""
        <div class="process-card">
            <div class="process-number">03</div>
            <div class="process-title">Kiểm tra</div>
            <div class="process-desc">Rà soát, chỉnh sửa và xác nhận dữ liệu</div>
        </div>
        """, unsafe_allow_html=True)

    with c4:
        st.markdown("""
        <div class="process-card">
            <div class="process-number">04</div>
            <div class="process-title">Xuất dữ liệu</div>
            <div class="process-desc">Tạo file Excel theo mẫu NSW chuẩn</div>
        </div>
        """, unsafe_allow_html=True)

# ===================== QUẢN LÝ HỒ SƠ =====================
elif menu == "🚢 Quản lý hồ sơ tàu":
    page_heading("Quản lý hồ sơ tàu", "Tạo mới và theo dõi thông tin các chuyến tàu nhập cảnh")
    st.info("Chức năng quản lý danh sách và thông tin hồ sơ tàu sẽ được phát triển ở bước tiếp theo.")

# ===================== TẢI LÊN =====================
elif menu == "📤 Tải lên chứng từ":
    page_heading("Tải lên chứng từ", "Tiếp nhận chứng từ phục vụ khai báo tàu nhập cảnh")

    uploaded_files = st.file_uploader(
        "Chọn chứng từ cần tải lên",
        type=["pdf", "png", "jpg", "jpeg", "xlsx", "docx"],
        accept_multiple_files=True
    )

    if uploaded_files:
        st.success(f"Đã tiếp nhận **{len(uploaded_files)}** tệp thành công.")
        st.subheader("Danh sách tệp đã tải lên")
        for file in uploaded_files:
            st.markdown(f"📄 **{file.name}** · {file.size/1024:.1f} KB")

# ===================== TRÍCH XUẤT =====================
elif menu == "🔍 Trích xuất dữ liệu":
    page_heading("Trích xuất dữ liệu", "Tự động nhận diện và trích xuất thông tin từ chứng từ")
    st.info("Chức năng OCR sẽ được tích hợp ở giai đoạn tiếp theo.")

# ===================== KIỂM TRA =====================
elif menu == "✅ Kiểm tra dữ liệu":
    page_heading("Kiểm tra dữ liệu", "Rà soát, chỉnh sửa và xác nhận thông tin được trích xuất")
    st.info("Khu vực kiểm tra và xác nhận dữ liệu sẽ được phát triển sau khi hoàn thiện chức năng OCR.")

# ===================== XUẤT EXCEL =====================
elif menu == "📊 Xuất Excel NSW":
    page_heading("Xuất dữ liệu Excel NSW", "Chuẩn hóa dữ liệu và xuất theo biểu mẫu khai báo NSW")
    st.info("Chức năng xuất Excel sẽ được tích hợp sau khi có Template NSW chính thức.")

# ===================== FOOTER =====================
st.markdown("---")
st.markdown("""
<div class="footer">
    SGCL · Hệ thống hỗ trợ tự động hóa nhập liệu chứng từ · Phiên bản thử nghiệm 1.0
</div>
""", unsafe_allow_html=True)
