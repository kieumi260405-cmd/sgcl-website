pip install pymupdf
import streamlit as st
import uuid
from datetime import datetime
import time
import hashlib
from pathlib import Path

st.set_page_config(
    page_title="SGCL – SHIPPING AGENT",
    page_icon="🚢",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# =========================================================
# 1. CẤU HÌNH LƯU TRỮ
# =========================================================

BASE_UPLOAD_DIR = Path("uploads")
BASE_UPLOAD_DIR.mkdir(parents=True, exist_ok=True)

ALLOWED_EXTENSIONS = {".pdf", ".xlsx", ".docx", ".png", ".jpg", ".jpeg"}
MAX_FILE_SIZE = 20 * 1024 * 1024


def get_ship_folder(ship):
    folder = BASE_UPLOAD_DIR / f"{ship['imo']}_{ship['id']}"
    folder.mkdir(parents=True, exist_ok=True)
    return folder


def calculate_file_hash(file_bytes):
    return hashlib.sha256(file_bytes).hexdigest()


def format_size(size):
    if size < 1024:
        return f"{size} B"
    elif size < 1024 * 1024:
        return f"{size / 1024:.1f} KB"
    return f"{size / (1024 * 1024):.2f} MB"


# =========================================================
# 2. XEM TRƯỚC CHỨNG TỪ
#    - PDF: chuyển từng trang thành ảnh bằng PyMuPDF
#      (Chrome chặn PDF nhúng dạng data:base64 trong iframe)
#    - DOCX: đọc trực tiếp cấu trúc XML, không cần python-docx
# =========================================================

@st.cache_data(show_spinner=False)
def render_pdf_pages(file_bytes, zoom=1.6):
    """Chuyển các trang PDF thành ảnh PNG."""
    import fitz  # PyMuPDF

    images = []
    with fitz.open(stream=file_bytes, filetype="pdf") as pdf:
        matrix = fitz.Matrix(zoom, zoom)
        for page in pdf:
            pix = page.get_pixmap(matrix=matrix)
            images.append(pix.tobytes("png"))
    return images


def read_docx_blocks(file_path):
    """
    Đọc nội dung DOCX theo đúng thứ tự (đoạn văn + bảng)
    bằng zipfile/XML, không phụ thuộc thư viện python-docx.
    """
    import zipfile
    import xml.etree.ElementTree as ET

    W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

    def paragraph_text(p):
        parts = []
        for node in p.iter():
            if node.tag == W + "t" and node.text:
                parts.append(node.text)
            elif node.tag == W + "tab":
                parts.append("\t")
            elif node.tag in (W + "br", W + "cr"):
                parts.append("\n")
        return "".join(parts)

    with zipfile.ZipFile(file_path) as z:
        root = ET.fromstring(z.read("word/document.xml"))

    body = root.find(W + "body")
    blocks = []

    for child in body:
        if child.tag == W + "p":
            text = paragraph_text(child)
            if text.strip():
                blocks.append(("text", text))

        elif child.tag == W + "tbl":
            rows = []
            for tr in child.findall(W + "tr"):
                row = []
                for tc in tr.findall(W + "tc"):
                    cell = "\n".join(
                        paragraph_text(p) for p in tc.findall(W + "p")
                    ).strip()
                    row.append(cell)
                rows.append(row)
            if rows:
                blocks.append(("table", rows))

    return blocks


def get_file_preview(file_path):
    """Xem trước chứng từ trực tiếp trên web."""
    import pandas as pd

    ext = file_path.suffix.lower()
    file_bytes = file_path.read_bytes()

    try:
        # ---------- PDF ----------
        if ext == ".pdf":
            st.download_button(
                "⬇️ Tải xuống PDF",
                data=file_bytes,
                file_name=file_path.name,
                mime="application/pdf",
                key=f"pdf_download_{file_path.name}"
            )

            try:
                with st.spinner("Đang hiển thị PDF..."):
                    pages = render_pdf_pages(file_bytes)
            except ImportError:
                st.warning(
                    "Thiếu thư viện PyMuPDF. Chạy lệnh: `pip install pymupdf` "
                    "rồi khởi động lại ứng dụng."
                )
                return

            total = len(pages)
            st.caption(f"Tổng số trang: {total}")

            if total == 0:
                st.warning("File PDF không có trang nào.")
                return

            if total > 1:
                view_mode = st.radio(
                    "Chế độ xem",
                    ["Từng trang", "Tất cả các trang"],
                    horizontal=True,
                    key=f"pdf_mode_{file_path.name}"
                )
            else:
                view_mode = "Tất cả các trang"

            if view_mode == "Từng trang":
                page_no = st.number_input(
                    "Trang",
                    min_value=1,
                    max_value=total,
                    value=1,
                    step=1,
                    key=f"pdf_page_{file_path.name}"
                )
                st.image(
                    pages[page_no - 1],
                    caption=f"Trang {page_no}/{total}",
                    use_container_width=True
                )
            else:
                for i, img in enumerate(pages, start=1):
                    st.image(
                        img,
                        caption=f"Trang {i}/{total}",
                        use_container_width=True
                    )

        # ---------- Ảnh ----------
        elif ext in [".png", ".jpg", ".jpeg"]:
            st.image(
                str(file_path),
                caption=file_path.name,
                use_container_width=True
            )

        # ---------- Excel ----------
        elif ext == ".xlsx":
            excel = pd.ExcelFile(file_path)

            sheet_name = st.selectbox(
                "Chọn trang tính",
                excel.sheet_names,
                key=f"preview_excel_{file_path.name}"
            )

            df = pd.read_excel(file_path, sheet_name=sheet_name)

            st.dataframe(
                df,
                use_container_width=True,
                hide_index=True
            )

        # ---------- Word ----------
        elif ext == ".docx":
            blocks = read_docx_blocks(file_path)

            if not blocks:
                st.info("File Word không có nội dung văn bản.")
                return

            with st.container(border=True):
                table_index = 0
                for kind, content in blocks:
                    if kind == "text":
                        st.markdown(
                            content.replace("\n", "  \n")
                        )
                    else:
                        table_index += 1
                        st.markdown(f"**Bảng {table_index}**")
                        width = max(len(r) for r in content)
                        padded = [r + [""] * (width - len(r)) for r in content]
                        df = pd.DataFrame(padded)
                        st.dataframe(
                            df,
                            use_container_width=True,
                            hide_index=True
                        )

        else:
            st.warning("Định dạng chưa hỗ trợ xem trước.")

    except Exception as e:
        st.error(f"Không thể xem trước chứng từ: {str(e)}")



# =========================================================
# 3. CSS – GIAO DIỆN TRẮNG VÀ XANH DƯƠNG
# =========================================================

css = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

html, body, .stApp, [data-testid="stAppViewContainer"],
[data-testid="stHeader"], .main, .block-container {
    background-color: #FFFFFF !important;
    color: #1E293B !important;
}

html, body, [class*="css"] {
    font-family: 'Inter', system-ui, sans-serif !important;
}

#MainMenu, footer, header {
    visibility: hidden !important;
}

.block-container {
    padding-top: 1.2rem !important;
    max-width: 1280px;
}

/* Nút chính */
.stButton > button[kind="primary"] {
    background-color: #2563EB !important;
    color: white !important;
    border: none !important;
    border-radius: 8px !important;
    font-weight: 600 !important;
}

.stButton > button[kind="primary"]:hover {
    background-color: #1D4ED8 !important;
}

/* Nút thường */
.stButton > button {
    background-color: #FFFFFF !important;
    color: #1E293B !important;
    border: 1px solid #2563EB !important;
    border-radius: 8px !important;
}

.stButton > button:hover {
    background-color: #EFF6FF !important;
}

/* Input */
.stTextInput input,
.stTextArea textarea {
    background-color: #FFFFFF !important;
    color: #1E293B !important;
    border: 1px solid #BFDBFE !important;
    border-radius: 8px !important;
}

.stTextInput input:focus {
    border-color: #2563EB !important;
    box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.15) !important;
}

/* Selectbox */
.stSelectbox div[data-baseweb="select"] > div {
    background-color: #FFFFFF !important;
    color: #1E293B !important;
    border: 1px solid #BFDBFE !important;
    border-radius: 8px !important;
}

/* File uploader */
[data-testid="stFileUploader"] {
    background-color: #F8FAFC !important;
    border: 2px dashed #93C5FD !important;
    border-radius: 12px !important;
    padding: 8px !important;
}

[data-testid="stFileUploader"] section {
    background-color: #F8FAFC !important;
    border: none !important;
}

[data-testid="stFileUploaderDropzone"] {
    background-color: #EFF6FF !important;
    border: 2px dashed #60A5FA !important;
    border-radius: 10px !important;
}

[data-testid="stFileUploader"] button {
    background-color: #FFFFFF !important;
    color: #1E40AF !important;
    border: 1px solid #93C5FD !important;
    border-radius: 6px !important;
}

[data-testid="stFileUploader"] p,
[data-testid="stFileUploader"] span,
[data-testid="stFileUploader"] small {
    color: #334155 !important;
}

/* Thông báo */
.stAlert,
div[data-testid="stAlert"] {
    background-color: #EFF6FF !important;
    color: #1E40AF !important;
    border: 1px solid #BFDBFE !important;
    border-radius: 10px !important;
}

.stAlert > div {
    background-color: #EFF6FF !important;
    color: #1E40AF !important;
}

/* Metric */
[data-testid="stMetric"] {
    background-color: #F8FAFC !important;
    border: 1px solid #E2E8F0 !important;
    border-radius: 10px !important;
    padding: 12px !important;
}

hr {
    border-color: #E2E8F0 !important;
}
</style>
"""

st.markdown(css, unsafe_allow_html=True)


# =========================================================
# 4. STATE
# =========================================================

if "ships" not in st.session_state:
    st.session_state.ships = []

if "page" not in st.session_state:
    st.session_state.page = "list"

if "current_ship_id" not in st.session_state:
    st.session_state.current_ship_id = None

if "show_form" not in st.session_state:
    st.session_state.show_form = False


def get_ship(sid):
    return next(
        (s for s in st.session_state.ships if s["id"] == sid),
        None
    )


def badge(status):
    styles = {
        "Nháp": ("#F1F5F9", "#334155"),
        "Đã tải chứng từ": ("#DBEAFE", "#1E40AF"),
        "Đã trích xuất": ("#BFDBFE", "#1D4ED8"),
        "Đã xác nhận": ("#93C5FD", "#1E3A8A"),
        "Hoàn thành": ("#2563EB", "#FFFFFF"),
    }

    bg, color = styles.get(status, ("#F1F5F9", "#334155"))
    return f'<span style="background:{bg};color:{color};padding:4px 12px;border-radius:9999px;font-size:12px;font-weight:500;">{status}</span>'


# =========================================================
# 5. HEADER
# =========================================================
st.markdown("""
<div style="display:flex;align-items:center;gap:12px;margin-bottom:24px;padding-bottom:16px;border-bottom:1px solid #DBEAFE;">
    <div style="width:42px;height:42px;background:#2563EB;border-radius:50%;display:flex;align-items:center;justify-content:center;color:white;font-size:20px;">🚢</div>
    <div>
        <div style="font-weight:700;font-size:18px;color:#000000;line-height:1.2;">SGCL</div>
        <div style="font-size:11px;color:#1D4ED8;letter-spacing:0.5px;">SHIPPING AGENT</div>
    </div>
</div>
""", unsafe_allow_html=True)


# =========================================================
# 6. TRANG DANH SÁCH HỒ SƠ TÀU
# =========================================================

if st.session_state.page == "list":

    col1, col2 = st.columns([3.5, 1.2])

    with col1:
        st.markdown("### 🚢 Hồ Sơ Khai Báo Tàu Nhập Cảnh")
        st.caption(
            "Quản lý các chuyến tàu nhập cảnh, "
            "tự động trích xuất chứng từ và xuất mẫu Excel"
        )

    with col2:
        if st.button(
            "➕ Tạo hồ sơ mới",
            type="primary",
            use_container_width=True
        ):
            st.session_state.show_form = True

    # ---------------- FORM TẠO HỒ SƠ ----------------

    if st.session_state.show_form:

        with st.form("form_ship", clear_on_submit=False):

            st.subheader("Tạo hồ sơ tàu mới")

            name = st.text_input("Tên tàu *", placeholder="VD: PACIFIC STAR")
            imo = st.text_input("Số IMO * (7 chữ số)", max_chars=7, placeholder="9438107")
            voyage = st.text_input("Số chuyến (Voyage No.) *", placeholder="V.2411N")
            call_sign = st.text_input("Hô hiệu (Call Sign) *", placeholder="3EAB9")

            c1, c2 = st.columns(2)

            with c1:
                submitted = st.form_submit_button("Hoàn tất", type="primary", use_container_width=True)

            with c2:
                cancel = st.form_submit_button("Hủy", use_container_width=True)

            if cancel:
                st.session_state.show_form = False
                st.rerun()

            if submitted:

                if not name.strip() or not imo.strip() or not voyage.strip() or not call_sign.strip():
                    st.error("Vui lòng điền đầy đủ thông tin.")

                elif not imo.isdigit() or len(imo) != 7:
                    st.error("Số IMO phải gồm đúng 7 chữ số.")

                elif any(s["imo"] == imo for s in st.session_state.ships):
                    st.error("Số IMO đã tồn tại.")

                else:
                    new_ship = {
                        "id": str(uuid.uuid4()),
                        "name": name.strip(),
                        "imo": imo,
                        "voyage_no": voyage.strip(),
                        "call_sign": call_sign.strip(),
                        "status": "Nháp",
                        "created_at": datetime.now().timestamp(),
                        "documents": [],
                        "extracted_data": None,
                        "step_status": {
                            "1": "done",
                            "2": "pending",
                            "3": "pending",
                            "4": "pending",
                            "5": "pending"
                        },
                        "current_step": 1,
                    }

                    st.session_state.ships.insert(0, new_ship)
                    st.session_state.show_form = False
                    st.success("Đã tạo hồ sơ thành công!")
                    time.sleep(0.5)
                    st.rerun()

    # ---------------- TÌM KIẾM VÀ LỌC ----------------

    c1, c2, c3 = st.columns([3, 1.3, 1.3])

    with c1:
        search = st.text_input(
            "Tìm kiếm",
            placeholder="Tìm theo tên tàu, số IMO, số chuyến, hô hiệu",
            label_visibility="collapsed"
        )

    with c2:
        status_f = st.selectbox(
            "Trạng thái",
            ["Tất cả trạng thái", "Nháp", "Đã tải chứng từ", "Đã trích xuất", "Đã xác nhận", "Hoàn thành"],
            label_visibility="collapsed"
        )

    with c3:
        sort_f = st.selectbox(
            "Sắp xếp",
            ["Mới tạo nhất trước", "Cũ nhất trước"],
            label_visibility="collapsed"
        )

    ships = st.session_state.ships[:]

    if search:
        q = search.lower()
        ships = [
            s for s in ships
            if q in s["name"].lower()
            or q in s["imo"]
            or q in s["voyage_no"].lower()
            or q in s["call_sign"].lower()
        ]

    if status_f != "Tất cả trạng thái":
        ships = [s for s in ships if s["status"] == status_f]

    ships = sorted(
        ships,
        key=lambda x: x["created_at"],
        reverse=(sort_f == "Mới tạo nhất trước")
    )

    st.markdown("")

    # ---------------- DANH SÁCH HỒ SƠ ----------------

    if not ships:
        st.info("Chưa có hồ sơ tàu nào. Nhấn **Tạo hồ sơ mới** để bắt đầu.")

    else:
        h1, h2, h3, h4, h5, h6 = st.columns([2.6, 1.3, 1.1, 1.5, 1.3, 1.6])

        h1.markdown("**Tên tàu / IMO**")
        h2.markdown("**Số chuyến**")
        h3.markdown("**Hô hiệu**")
        h4.markdown("**Trạng thái**")
        h5.markdown("**Ngày tạo**")
        h6.markdown("**Thao tác**")

        st.divider()

        for ship in ships:
            c1, c2, c3, c4, c5, c6 = st.columns([2.6, 1.3, 1.1, 1.5, 1.3, 1.6])

            with c1:
                st.markdown(f"**{ship['name']}**  `IMO: {ship['imo']}`")

            with c2:
                st.write(ship["voyage_no"])

            with c3:
                st.write(ship["call_sign"])

            with c4:
                st.markdown(badge(ship["status"]), unsafe_allow_html=True)

            with c5:
                st.caption(datetime.fromtimestamp(ship["created_at"]).strftime("%d/%m/%Y %H:%M"))

            with c6:
                b1, b2 = st.columns(2)

                if b1.button("Xử lý", key=f"p_{ship['id']}", use_container_width=True):
                    st.session_state.current_ship_id = ship["id"]
                    st.session_state.page = "detail"
                    st.rerun()

                if b2.button("🗑️", key=f"d_{ship['id']}", use_container_width=True):
                    folder = BASE_UPLOAD_DIR / f"{ship['imo']}_{ship['id']}"
                    if folder.exists():
                        import shutil
                        shutil.rmtree(folder, ignore_errors=True)

                    st.session_state.ships = [s for s in st.session_state.ships if s["id"] != ship["id"]]
                    st.rerun()

            st.divider()


# =========================================================
# 7. TRANG CHI TIẾT HỒ SƠ TÀU
# =========================================================

else:

    ship = get_ship(st.session_state.current_ship_id)

    if not ship:
        st.session_state.page = "list"
        st.rerun()

    if st.button("← Danh sách hồ sơ"):
        st.session_state.page = "list"
        st.rerun()

    st.markdown(f"### {ship['name']} `(IMO: {ship['imo']})`")

    st.markdown(f"""
    <div style="background:#FFFFFF;border:1px solid #BFDBFE;border-radius:14px;padding:20px;margin:12px 0 20px 0;">
        <div style="display:flex;gap:16px;align-items:center;">
            <div style="width:52px;height:52px;background:#DBEAFE;border-radius:12px;display:flex;align-items:center;justify-content:center;font-size:26px;">
                🚢
            </div>
            <div>
                <div style="font-size:18px;font-weight:700;color:#000000;">{ship['name']}</div>
                <div style="color:#1D4ED8;font-size:13px;margin-top:2px;">Thông tin hồ sơ khai báo tàu nhập cảnh</div>
                <div style="margin-top:10px;display:flex;gap:28px;font-size:14px;color:#000000;">
                    <span>SỐ IMO: <b>{ship['imo']}</b></span>
                    <span>SỐ CHUYẾN: <b>{ship['voyage_no']}</b></span>
                    <span>HÔ HIỆU: <b>{ship['call_sign']}</b></span>
                </div>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # ---------------- THANH 5 BƯỚC ----------------

    steps = [
        (1, "Thông tin tàu"),
        (2, "Tải chứng từ"),
        (3, "Trích xuất"),
        (4, "Kiểm tra"),
        (5, "Xuất Excel")
    ]

    cols = st.columns(5)

    for i, (sid, title) in enumerate(steps):
        active = ship["current_step"] == sid
        done = ship["step_status"].get(str(sid)) == "done"

        label = f"{'✅' if done else '🔵' if active else '⚪'} {title}"

        if cols[i].button(label, key=f"st_{sid}", use_container_width=True):
            ship["current_step"] = sid
            st.rerun()

    st.markdown("---")

    step = ship["current_step"]


    # =====================================================
    # BƯỚC 1 – THÔNG TIN TÀU
    # =====================================================

    if step == 1:

        st.subheader("1. Thông tin tàu")

        c1, c2 = st.columns(2)

        c1.info(f"**Tên tàu:** {ship['name']}")
        c1.info(f"**Số IMO:** {ship['imo']}")

        c2.info(f"**Số chuyến:** {ship['voyage_no']}")
        c2.info(f"**Hô hiệu:** {ship['call_sign']}")

        if st.button("Xác nhận & Tiếp tục →", type="primary"):
            ship["step_status"]["1"] = "done"
            ship["current_step"] = 2
            st.rerun()


    # =====================================================
    # BƯỚC 2 – TẢI CHỨNG TỪ
    # =====================================================

    elif step == 2:

        st.subheader("2. Tải chứng từ")

        st.caption("Tiếp nhận và quản lý chứng từ khai báo tàu nhập cảnh.")

        ship_folder = get_ship_folder(ship)

        if "documents" not in ship:
            ship["documents"] = []

        total_files = len(ship["documents"])
        total_size = sum(d.get("size", 0) for d in ship["documents"])

        c1, c2, c3 = st.columns(3)

        c1.metric("Tổng chứng từ", total_files)
        c2.metric("Dung lượng", format_size(total_size))
        c3.metric("Định dạng hỗ trợ", "PDF, XLSX, DOCX, JPG, PNG")

        st.divider()

        st.markdown("#### 📤 Tải lên chứng từ")

        uploaded_files = st.file_uploader(
            "Chọn một hoặc nhiều chứng từ",
            type=["pdf", "xlsx", "docx", "png", "jpg", "jpeg"],
            accept_multiple_files=True,
            key=f"upload_{ship['id']}"
        )

        if uploaded_files:

            if st.button("⬆️ Xác nhận tải lên", type="primary", key=f"confirm_upload_{ship['id']}"):

                success_count = 0
                errors = []

                existing_hashes = {d.get("hash") for d in ship["documents"] if d.get("hash")}
                existing_names = {d["name"].lower() for d in ship["documents"]}

                for uploaded_file in uploaded_files:

                    file_name = Path(uploaded_file.name).name
                    file_ext = Path(file_name).suffix.lower()
                    file_bytes = uploaded_file.getvalue()

                    if file_ext not in ALLOWED_EXTENSIONS:
                        errors.append(f"{file_name}: Định dạng không được hỗ trợ.")
                        continue

                    if not file_bytes:
                        errors.append(f"{file_name}: File rỗng hoặc không hợp lệ.")
                        continue

                    if len(file_bytes) > MAX_FILE_SIZE:
                        errors.append(f"{file_name}: File vượt quá giới hạn 20 MB.")
                        continue

                    if file_name.lower() in existing_names:
                        errors.append(f"{file_name}: Tên file đã tồn tại trong hồ sơ.")
                        continue

                    file_hash = calculate_file_hash(file_bytes)

                    if file_hash in existing_hashes:
                        errors.append(f"{file_name}: Nội dung file đã tồn tại.")
                        continue

                    file_path = ship_folder / file_name

                    try:
                        file_path.write_bytes(file_bytes)

                        ship["documents"].append({
                            "name": file_name,
                            "size": len(file_bytes),
                            "hash": file_hash,
                            "path": str(file_path),
                            "uploaded_at": datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                        })

                        existing_names.add(file_name.lower())
                        existing_hashes.add(file_hash)
                        success_count += 1

                    except OSError as e:
                        errors.append(f"{file_name}: Không thể lưu file ({e}).")

                if success_count > 0:
                    ship["status"] = "Đã tải chứng từ"
                    st.success(f"Đã tải thành công {success_count} chứng từ.")

                if errors:
                    for error in errors:
                        st.warning(error)

                st.rerun()

        st.divider()

        st.markdown("#### 📁 Danh sách chứng từ")

        if not ship["documents"]:

            st.markdown("""
            <div style="background:#EFF6FF;border:1px solid #BFDBFE;border-radius:12px;padding:18px 22px;color:#1E40AF;font-size:15px;display:flex;align-items:center;gap:12px;margin-bottom:12px;">
                <span style="font-size:22px;">📂</span>
                <div>
                    <strong>Hồ sơ chưa có chứng từ</strong><br>
                    <span style="font-size:13.5px;opacity:0.85;">Vui lòng tải file lên để tiếp tục xử lý.</span>
                </div>
            </div>
            """, unsafe_allow_html=True)

        else:

            st.write(f"Đang quản lý **{len(ship['documents'])} chứng từ**")

            for i, doc in enumerate(ship["documents"]):

                file_path = Path(doc.get("path", ""))
                file_exists = file_path.is_file()

                with st.container(border=True):

                    col1, col2, col3, col4 = st.columns([3.5, 1.2, 1.2, 1.1])

                    with col1:
                        ext = Path(doc["name"]).suffix.upper()
                        st.markdown(f"**📄 {doc['name']}**")
                        st.caption(f"{ext} · {format_size(doc['size'])} · {doc.get('uploaded_at', 'Không rõ thời gian')}")

                        if not file_exists:
                            st.error("File không tồn tại trên hệ thống.")

                    with col2:
                        if file_exists:
                            with open(file_path, "rb") as f:
                                file_data = f.read()

                            st.download_button(
                                "⬇️ Tải xuống",
                                data=file_data,
                                file_name=doc["name"],
                                key=f"download_{ship['id']}_{i}",
                                use_container_width=True
                            )

                    with col3:
                        if file_exists:
                            if st.button("👁️ Xem", key=f"preview_{ship['id']}_{i}", use_container_width=True):
                                st.session_state[f"preview_doc_{ship['id']}"] = doc["name"]

                    with col4:
                        if st.button("🗑️ Xóa", key=f"delete_{ship['id']}_{i}", use_container_width=True):
                            if file_exists:
                                file_path.unlink()

                            preview_key = f"preview_doc_{ship['id']}"
                            if st.session_state.get(preview_key) == doc["name"]:
                                del st.session_state[preview_key]

                            ship["documents"].pop(i)

                            if not ship["documents"]:
                                ship["status"] = "Nháp"

                            st.success("Đã xóa chứng từ.")
                            st.rerun()

            # ---------------- XEM TRƯỚC ----------------

            preview_name = st.session_state.get(f"preview_doc_{ship['id']}")

            if preview_name:
                preview_doc = next((d for d in ship["documents"] if d["name"] == preview_name), None)

                if preview_doc:
                    preview_path = Path(preview_doc["path"])

                    st.divider()
                    st.markdown("#### 👁️ Xem trước chứng từ")
                    st.markdown(f"**{preview_name}**")

                    if preview_path.is_file():
                        get_file_preview(preview_path)
                    else:
                        st.error("Không tìm thấy file để xem trước.")

                    if st.button("Đóng xem trước", key=f"close_preview_{ship['id']}"):
                        del st.session_state[f"preview_doc_{ship['id']}"]
                        st.rerun()

        # ---------------- ĐIỀU HƯỚNG ----------------

        st.divider()

        c1, c2 = st.columns(2)

        with c1:
            if st.button("← Quay lại", use_container_width=True):
                ship["current_step"] = 1
                st.rerun()

        with c2:
            if st.button(
                "Hoàn tất & Tiếp tục →",
                type="primary",
                use_container_width=True,
                disabled=not ship["documents"]
            ):
                ship["step_status"]["2"] = "done"
                ship["current_step"] = 3
                st.rerun()


    # =====================================================
    # BƯỚC 3 – OCR
    # =====================================================

    elif step == 3:

        st.subheader("3. Trích xuất dữ liệu (OCR)")

        if not ship["documents"]:
            st.warning("Chưa có chứng từ. Vui lòng quay lại bước 2.")
        else:
            st.info(
                "Chức năng OCR đang chờ cấu hình. "
                "Các trường thông tin sẽ được thiết lập "
                "theo danh sách chứng từ và yêu cầu trích xuất."
            )

            st.markdown("""
            <div style="background:#F8FAFC;border:1px solid #E2E8F0;border-radius:12px;padding:20px;">
                <div style="font-size:16px;font-weight:600;color:#1E293B;">🤖 Hệ thống trích xuất dữ liệu</div>
                <p style="color:#64748B;font-size:14px;margin-top:8px;">
                    Chưa có dữ liệu OCR. Hệ thống sẽ được cấu hình sau khi xác định các trường cần đọc từ từng loại chứng từ.
                </p>
                <div style="color:#2563EB;font-size:13px;font-weight:500;">Trạng thái: Chờ cấu hình OCR</div>
            </div>
            """, unsafe_allow_html=True)

        c1, c2 = st.columns(2)

        with c1:
            if st.button("← Quay lại"):
                ship["current_step"] = 2
                st.rerun()

        with c2:
            st.button("Xác nhận & Tiếp tục →", type="primary", disabled=True, use_container_width=True)


    # =====================================================
    # BƯỚC 4 – KIỂM TRA DỮ LIỆU
    # =====================================================

    elif step == 4:

        st.subheader("4. Kiểm tra dữ liệu")

        st.info(
            "Bước kiểm tra sẽ được kích hoạt sau khi hoàn thiện "
            "chức năng OCR và cấu trúc dữ liệu trích xuất."
        )

        if st.button("← Quay lại"):
            ship["current_step"] = 3
            st.rerun()


    # =====================================================
    # BƯỚC 5 – XUẤT EXCEL
    # =====================================================

    elif step == 5:

        st.subheader("5. Xuất Excel NSW")

        st.info(
            "Chức năng xuất Excel sẽ được hoàn thiện sau khi "
            "xác định cấu trúc dữ liệu OCR và mẫu Excel NSW."
        )

        if st.button("← Quay lại"):
            ship["current_step"] = 4
            st.rerun()
