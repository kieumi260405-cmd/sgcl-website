
import streamlit as st
import uuid
from datetime import datetime
import time
import hashlib
from pathlib import Path
import os
import json
import base64
from io import BytesIO
from PIL import Image
from openai import OpenAI
from streamlit_drawable_canvas import st_canvas
import fitz
import openpyxl
import sys
import os
import locale

os.environ["PYTHONIOENCODING"] = "utf-8"
os.environ["LANG"] = "C.UTF-8"
os.environ["LC_ALL"] = "C.UTF-8"
for stream in (sys.stdout, sys.stderr):
    if hasattr(stream, "reconfigure"):
        stream.reconfigure(
            encoding="utf-8",
            errors="replace"
        )
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
ALLOWED_EXTENSIONS = {".pdf", ".png", ".jpg", ".jpeg"}
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
# =========================================================

@st.cache_data(show_spinner=False)
def render_pdf_pages(file_bytes, zoom=1.6):
    """Chuyển các trang PDF thành ảnh PNG."""
    images = []

    with fitz.open(stream=file_bytes, filetype="pdf") as pdf:
        matrix = fitz.Matrix(zoom, zoom)

        for page in pdf:
            pix = page.get_pixmap(matrix=matrix)
            images.append(pix.tobytes("png"))

    return images


def get_file_preview(file_path):
    """Xem trước chứng từ trực tiếp trên web."""
    ext = file_path.suffix.lower()
    file_bytes = file_path.read_bytes()

    try:
        if ext == ".pdf":
            st.download_button(
                "⬇️ Tải xuống PDF",
                data=file_bytes,
                file_name=file_path.name,
                mime="application/pdf",
                key=f"pdf_download_{file_path.name}"
            )

            with st.spinner("Đang hiển thị PDF..."):
                pages = render_pdf_pages(file_bytes)

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

        elif ext in [".png", ".jpg", ".jpeg"]:
            st.image(
                str(file_path),
                caption=file_path.name,
                use_container_width=True
            )

        else:
            st.warning("Định dạng chưa hỗ trợ xem trước.")

    except Exception as e:
        st.error(f"Không thể xem trước chứng từ: {str(e)}")


# =========================================================
# 3. CÁC HÀM HỖ TRỢ AI - BƯỚC 3
# =========================================================

def sgcl_get_api_key():
    """Lấy API Key từ Streamlit Secrets hoặc Google Colab."""

    api_key = None

    try:
        api_key = st.secrets.get("OPENAI_API_KEY")
    except Exception:
        api_key = None

    if not api_key:
        api_key = os.environ.get("OPENAI_API_KEY")

    if not api_key:
        raise ValueError(
            "Chưa cấu hình OPENAI_API_KEY. "
            "Vui lòng kiểm tra phần Secrets."
        )

    return str(api_key).strip()
def sgcl_get_document_bytes(doc):
    """Đọc chứng từ đã tải lên ở Bước 2."""
    file_path = Path(doc["path"])

    if not file_path.exists():
        raise FileNotFoundError(
            f"Không tìm thấy chứng từ: {doc['name']}"
        )

    return file_path.read_bytes()


def sgcl_render_pages(file_bytes, file_name):
    """Chuyển PDF hoặc ảnh thành danh sách ảnh để AI đọc."""
    ext = Path(file_name).suffix.lower()
    images = []

    if ext == ".pdf":
        pdf = fitz.open(stream=file_bytes, filetype="pdf")

        try:
            for page in pdf:
                pix = page.get_pixmap(
                    matrix=fitz.Matrix(2, 2),
                    alpha=False
                )

                image = Image.open(
                    BytesIO(pix.tobytes("png"))
                ).convert("RGB")

                images.append(image)
        finally:
            pdf.close()

    elif ext in [".png", ".jpg", ".jpeg"]:
        image = Image.open(
            BytesIO(file_bytes)
        ).convert("RGB")

        images.append(image)

    else:
        raise ValueError(
            "Bước 3 chỉ hỗ trợ trích xuất dữ liệu từ PDF và hình ảnh."
        )

    return images


def sgcl_image_to_data_url(image):
    """Chuyển ảnh thành dữ liệu để gửi đến AI."""
    buffer = BytesIO()
    image.save(buffer, format="PNG")

    image_base64 = base64.b64encode(
        buffer.getvalue()
    ).decode("utf-8")

    return f"data:image/png;base64,{image_base64}"


def sgcl_extract_with_ai(
    image,
    field_name,
    extraction_type="single"
):
    """Gửi vùng ảnh hoặc toàn trang đến AI."""

    client = OpenAI(
        api_key=sgcl_get_api_key()
    )

    image_url = sgcl_image_to_data_url(image)

    if extraction_type == "table":
        output_format = """
        Trả về JSON theo cấu trúc:
        {
            "headers": ["Tên cột 1", "Tên cột 2"],
            "rows": [
                ["Giá trị hàng 1", "Giá trị hàng 1"],
                ["Giá trị hàng 2", "Giá trị hàng 2"]
            ]
        }
        Nếu không đọc được, trả về headers và rows rỗng.
        """
    else:
        output_format = """
        Trả về JSON theo cấu trúc:
        {
            "value": "Giá trị được trích xuất"
        }
        Nếu không đọc được, value là chuỗi rỗng.
        """

    prompt = f"""
    Bạn là hệ thống AI hỗ trợ xử lý chứng từ hàng hải.

    Yêu cầu:
    Chỉ trích xuất thông tin sau: {field_name}

    Không tự suy đoán hoặc tự bổ sung dữ liệu.
    Nếu thông tin không rõ, hãy để trống.
    Giữ nguyên nội dung, số liệu và định dạng quan trọng
    như ngày tháng, số hiệu, tên riêng.

    {output_format}

    Chỉ trả về JSON hợp lệ, không giải thích thêm.
    """

    response = client.responses.create(
        model=os.environ.get(
            "OPENAI_MODEL",
            "gpt-4.1-mini"
        ),
        input=[
            {
                "role": "user",
                "content": [
                    {
                        "type": "input_text",
                        "text": prompt
                    },
                    {
                        "type": "input_image",
                        "image_url": image_url
                    }
                ]
            }
        ]
    )

    raw_result = str(
    response.output_text or ""
).strip()

    if raw_result.startswith("```"):
        raw_result = raw_result.replace(
            "```json", ""
        ).replace("```", "").strip()

    return json.loads(raw_result)


def sgcl_write_to_excel(
    workbook_bytes,
    sheet_name,
    start_cell,
    result,
    extraction_type="single"
):
    """Ghi kết quả AI vào đúng sheet và ô Excel được chọn."""

    workbook = openpyxl.load_workbook(
        BytesIO(workbook_bytes)
    )

    if sheet_name not in workbook.sheetnames:
        workbook.close()
        raise ValueError(
            f"Không tìm thấy sheet: {sheet_name}"
        )

    ws = workbook[sheet_name]

    if extraction_type == "table":
        headers = result.get("headers", [])
        rows = result.get("rows", [])

        if not headers and not rows:
            workbook.close()
            raise ValueError(
                "AI không trích xuất được dữ liệu bảng."
            )

        start = openpyxl.utils.cell.coordinate_from_string(
            start_cell.upper()
        )

        start_col = openpyxl.utils.column_index_from_string(
            start[0]
        )

        start_row = start[1]

        for col_index, header in enumerate(headers):
            ws.cell(
                row=start_row,
                column=start_col + col_index,
                value=header
            )

        for row_index, row_data in enumerate(rows, start=1):
            for col_index, value in enumerate(row_data):
                ws.cell(
                    row=start_row + row_index,
                    column=start_col + col_index,
                    value=value
                )

    else:
        ws[start_cell.upper()] = result.get("value", "")

    output = BytesIO()
    workbook.save(output)
    workbook.close()

    return output.getvalue()


# =========================================================
# 4. DANH SÁCH SHEET NSW
# =========================================================

NSW_SHEETS = [
    "Notice ship to port",
    "Confirmation of Arrival",
    "General Declaration",
    "Goods Declaration",
    "Crew list",
    "Crew’s Effects Declaration",
    "Ship’s store declaration",
    "Declaration For Ship Security",
    "Declaration For Plant",
    "Declaration For Health"
]


# =========================================================
# 5. CSS – GIAO DIỆN TRẮNG VÀ XANH DƯƠNG
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
# 6. STATE
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

    return (
        f'<span style="background:{bg};color:{color};'
        f'padding:4px 12px;border-radius:9999px;'
        f'font-size:12px;font-weight:500;">{status}</span>'
    )


# =========================================================
# 7. HEADER
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
# 8. TRANG DANH SÁCH HỒ SƠ TÀU
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

    # FORM TẠO HỒ SƠ

    if st.session_state.show_form:

        with st.form("form_ship", clear_on_submit=False):

            st.subheader("Tạo hồ sơ tàu mới")

            name = st.text_input(
                "Tên tàu *",
                placeholder="VD: PACIFIC STAR"
            )

            imo = st.text_input(
                "Số IMO * (7 chữ số)",
                max_chars=7,
                placeholder="9438107"
            )

            voyage = st.text_input(
                "Số chuyến (Voyage No.) *",
                placeholder="V.2411N"
            )

            call_sign = st.text_input(
                "Hô hiệu (Call Sign) *",
                placeholder="3EAB9"
            )

            c1, c2 = st.columns(2)

            with c1:
                submitted = st.form_submit_button(
                    "Hoàn tất",
                    type="primary",
                    use_container_width=True
                )

            with c2:
                cancel = st.form_submit_button(
                    "Hủy",
                    use_container_width=True
                )

            if cancel:
                st.session_state.show_form = False
                st.rerun()

            if submitted:

                if (
                    not name.strip()
                    or not imo.strip()
                    or not voyage.strip()
                    or not call_sign.strip()
                ):
                    st.error("Vui lòng điền đầy đủ thông tin.")

                elif not imo.isdigit() or len(imo) != 7:
                    st.error("Số IMO phải gồm đúng 7 chữ số.")

                elif any(
                    s["imo"] == imo
                    for s in st.session_state.ships
                ):
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

    # TÌM KIẾM VÀ LỌC

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
            [
                "Tất cả trạng thái",
                "Nháp",
                "Đã tải chứng từ",
                "Đã trích xuất",
                "Đã xác nhận",
                "Hoàn thành"
            ],
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
        ships = [
            s for s in ships
            if s["status"] == status_f
        ]

    ships = sorted(
        ships,
        key=lambda x: x["created_at"],
        reverse=(sort_f == "Mới tạo nhất trước")
    )

    st.markdown("")

    # DANH SÁCH HỒ SƠ

    if not ships:
        st.info(
            "Chưa có hồ sơ tàu nào. "
            "Nhấn **Tạo hồ sơ mới** để bắt đầu."
        )

    else:
        h1, h2, h3, h4, h5, h6 = st.columns(
            [2.6, 1.3, 1.1, 1.5, 1.3, 1.6]
        )

        h1.markdown("**Tên tàu / IMO**")
        h2.markdown("**Số chuyến**")
        h3.markdown("**Hô hiệu**")
        h4.markdown("**Trạng thái**")
        h5.markdown("**Ngày tạo**")
        h6.markdown("**Thao tác**")

        st.divider()

        for ship in ships:

            c1, c2, c3, c4, c5, c6 = st.columns(
                [2.6, 1.3, 1.1, 1.5, 1.3, 1.6]
            )

            with c1:
                st.markdown(
                    f"**{ship['name']}**  `IMO: {ship['imo']}`"
                )

            with c2:
                st.write(ship["voyage_no"])

            with c3:
                st.write(ship["call_sign"])

            with c4:
                st.markdown(
                    badge(ship["status"]),
                    unsafe_allow_html=True
                )

            with c5:
                st.caption(
                    datetime.fromtimestamp(
                        ship["created_at"]
                    ).strftime("%d/%m/%Y %H:%M")
                )

            with c6:
                b1, b2 = st.columns(2)

                if b1.button(
                    "Xử lý",
                    key=f"p_{ship['id']}",
                    use_container_width=True
                ):
                    st.session_state.current_ship_id = ship["id"]
                    st.session_state.page = "detail"
                    st.rerun()

                if b2.button(
                    "🗑️",
                    key=f"d_{ship['id']}",
                    use_container_width=True
                ):
                    folder = (
                        BASE_UPLOAD_DIR
                        / f"{ship['imo']}_{ship['id']}"
                    )

                    if folder.exists():
                        import shutil
                        shutil.rmtree(folder, ignore_errors=True)

                    st.session_state.ships = [
                        s for s in st.session_state.ships
                        if s["id"] != ship["id"]
                    ]

                    st.rerun()

            st.divider()


# =========================================================
# 9. TRANG CHI TIẾT HỒ SƠ TÀU
# =========================================================

else:

    ship = get_ship(st.session_state.current_ship_id)

    if not ship:
        st.session_state.page = "list"
        st.rerun()

    if st.button("← Danh sách hồ sơ"):
        st.session_state.page = "list"
        st.rerun()

    st.markdown(
        f"### {ship['name']} `(IMO: {ship['imo']})`"
    )

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

    # THANH 5 BƯỚC

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

        label = (
            f"{'✅' if done else '🔵' if active else '⚪'} {title}"
        )

        if cols[i].button(
            label,
            key=f"st_{sid}",
            use_container_width=True
        ):
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

        if st.button(
            "Xác nhận & Tiếp tục →",
            type="primary"
        ):
            ship["step_status"]["1"] = "done"
            ship["current_step"] = 2
            st.rerun()

    # =====================================================
    # BƯỚC 2 – TẢI CHỨNG TỪ
    # =====================================================

    elif step == 2:

        st.subheader("2. Tải chứng từ")

        st.caption(
            "Tiếp nhận và quản lý chứng từ khai báo tàu nhập cảnh."
        )

        ship_folder = get_ship_folder(ship)

        if "documents" not in ship:
            ship["documents"] = []

        total_files = len(ship["documents"])
        total_size = sum(
            d.get("size", 0)
            for d in ship["documents"]
        )

        c1, c2, c3 = st.columns(3)

        c1.metric("Tổng chứng từ", total_files)
        c2.metric("Dung lượng", format_size(total_size))
        c3.metric("Định dạng hỗ trợ", "PDF, JPG, PNG")

        st.divider()

        st.markdown("#### 📤 Tải lên chứng từ")

        uploaded_files = st.file_uploader(
            "Chọn một hoặc nhiều chứng từ",
            type=["pdf", "png", "jpg", "jpeg"],
            accept_multiple_files=True,
            key=f"upload_{ship['id']}"
        )

        if uploaded_files:

            if st.button(
                "⬆️ Xác nhận tải lên",
                type="primary",
                key=f"confirm_upload_{ship['id']}"
            ):

                success_count = 0
                errors = []

                existing_hashes = {
                    d.get("hash")
                    for d in ship["documents"]
                    if d.get("hash")
                }

                existing_names = {
                    d["name"].lower()
                    for d in ship["documents"]
                }

                for uploaded_file in uploaded_files:

                    file_name = Path(uploaded_file.name).name
                    file_ext = Path(file_name).suffix.lower()
                    file_bytes = uploaded_file.getvalue()

                    if file_ext not in ALLOWED_EXTENSIONS:
                        errors.append(
                            f"{file_name}: Định dạng không được hỗ trợ."
                        )
                        continue

                    if not file_bytes:
                        errors.append(
                            f"{file_name}: File rỗng hoặc không hợp lệ."
                        )
                        continue

                    if len(file_bytes) > MAX_FILE_SIZE:
                        errors.append(
                            f"{file_name}: File vượt quá giới hạn 20 MB."
                        )
                        continue

                    if file_name.lower() in existing_names:
                        errors.append(
                            f"{file_name}: Tên file đã tồn tại trong hồ sơ."
                        )
                        continue

                    file_hash = calculate_file_hash(file_bytes)

                    if file_hash in existing_hashes:
                        errors.append(
                            f"{file_name}: Nội dung file đã tồn tại."
                        )
                        continue

                    file_path = ship_folder / file_name

                    try:
                        file_path.write_bytes(file_bytes)

                        ship["documents"].append({
                            "name": file_name,
                            "size": len(file_bytes),
                            "hash": file_hash,
                            "path": str(file_path),
                            "uploaded_at": datetime.now().strftime(
                                "%d/%m/%Y %H:%M:%S"
                            )
                        })

                        existing_names.add(file_name.lower())
                        existing_hashes.add(file_hash)
                        success_count += 1

                    except OSError as e:
                        errors.append(
                            f"{file_name}: Không thể lưu file ({e})."
                        )

                if success_count > 0:
                    ship["status"] = "Đã tải chứng từ"
                    st.success(
                        f"Đã tải thành công {success_count} chứng từ."
                    )

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

            st.write(
                f"Đang quản lý **{len(ship['documents'])} chứng từ**"
            )

            for i, doc in enumerate(ship["documents"]):

                file_path = Path(doc.get("path", ""))
                file_exists = file_path.is_file()

                with st.container(border=True):

                    col1, col2, col3, col4 = st.columns(
                        [3.5, 1.2, 1.2, 1.1]
                    )

                    with col1:
                        ext = Path(doc["name"]).suffix.upper()

                        st.markdown(f"**📄 {doc['name']}**")

                        st.caption(
                            f"{ext} · {format_size(doc['size'])} · "
                            f"{doc.get('uploaded_at', 'Không rõ thời gian')}"
                        )

                        if not file_exists:
                            st.error(
                                "File không tồn tại trên hệ thống."
                            )

                   
                    with col2:
                        if file_exists:
                            file_data = file_path.read_bytes()

                            st.download_button(
                                "⬇️ Tải xuống",
                                data=file_data,
                                file_name=doc["name"],
                                mime="application/octet-stream",
                                key=f"download_{ship['id']}_{i}",
                                use_container_width=True
                            )

                    with col3:
                        if file_exists:
                            if st.button(
                                "👁️ Xem",
                                key=f"preview_{ship['id']}_{i}",
                                use_container_width=True
                            ):
                                st.session_state[
                                    f"preview_doc_{ship['id']}"
                                ] = doc["name"]

                    with col4:
                        if st.button(
                            "🗑️ Xóa",
                            key=f"delete_{ship['id']}_{i}",
                            use_container_width=True
                        ):
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

            # XEM TRƯỚC

            preview_name = st.session_state.get(
                f"preview_doc_{ship['id']}"
            )

            if preview_name:

                preview_doc = next(
                    (
                        d for d in ship["documents"]
                        if d["name"] == preview_name
                    ),
                    None
                )

                if preview_doc:

                    preview_path = Path(preview_doc["path"])

                    st.divider()
                    st.markdown("#### 👁️ Xem trước chứng từ")
                    st.markdown(f"**{preview_name}**")

                    if preview_path.is_file():
                        get_file_preview(preview_path)
                    else:
                        st.error(
                            "Không tìm thấy file để xem trước."
                        )

                    if st.button(
                        "Đóng xem trước",
                        key=f"close_preview_{ship['id']}"
                    ):
                        del st.session_state[
                            f"preview_doc_{ship['id']}"
                        ]
                        st.rerun()

        # ĐIỀU HƯỚNG

        st.divider()

        c1, c2 = st.columns(2)

        with c1:
            if st.button(
                "← Quay lại",
                use_container_width=True
            ):
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
    # BƯỚC 3 – TRÍCH XUẤT DỮ LIỆU BẰNG AI
    # =====================================================

    elif step == 3:

        st.subheader("3. Trích xuất dữ liệu bằng AI")

        st.caption(
            "AI đọc chứng từ và trích xuất đúng những trường "
            "thông tin do người dùng yêu cầu."
        )

        ship.setdefault("field_extractions", [])

        if not ship["documents"]:

            st.warning(
                "Chưa có chứng từ. Vui lòng quay lại Bước 2."
            )

        else:

            st.markdown("#### 1. Chọn chứng từ")

            doc_index = st.selectbox(
                "Chứng từ cần xử lý",
                range(len(ship["documents"])),
                format_func=lambda i: ship["documents"][i]["name"],
                key=f"ai_doc_{ship['id']}"
            )

            selected_doc = ship["documents"][doc_index]

            try:

                document_bytes = sgcl_get_document_bytes(
                    selected_doc
                )

                pages = sgcl_render_pages(
                    document_bytes,
                    selected_doc["name"]
                )

                # ĐÃ SỬA LỖI THỤT LỀ TẠI ĐÂY
                if not pages:
                    raise ValueError(
                        "Chứng từ không có trang dữ liệu để trích xuất."
                    )

                st.markdown("#### 2. Chọn trang cần trích xuất")

                page_number = st.selectbox(
                    "Trang chứng từ",
                    range(len(pages)),
                    format_func=lambda i: f"Trang {i + 1}",
                    key=f"ai_page_{ship['id']}_{doc_index}"
                )

                original_image = pages[page_number]

                st.markdown("#### 3. Xác định vùng dữ liệu")

                extraction_mode = st.radio(
                    "Phương thức trích xuất",
                    [
                        "Đọc toàn bộ trang",
                        "Khoanh vùng thông tin cần đọc"
                    ],
                    horizontal=True,
                    key=f"ai_mode_{ship['id']}"
                )

                selected_image = original_image
                region_selected = False

                if extraction_mode == "Đọc toàn bộ trang":

                    st.image(
                        original_image,
                        caption="Toàn bộ trang chứng từ",
                        use_container_width=True
                    )

                else:

                    st.info(
                        "Kéo chuột để tạo hình chữ nhật bao quanh "
                        "vùng thông tin muốn trích xuất."
                    )

                    display_image = original_image.copy()
                    display_image.thumbnail((900, 1200))

                    canvas_width, canvas_height = display_image.size

                    canvas_result = st_canvas(
                        fill_color="rgba(37, 99, 235, 0.15)",
                        stroke_width=2,
                        stroke_color="#2563EB",
                        background_image=display_image,
                        update_streamlit=True,
                        height=canvas_height,
                        width=canvas_width,
                        drawing_mode="rect",
                        key=(
                            f"canvas_{ship['id']}_"
                            f"{doc_index}_{page_number}"
                        )
                    )

                    if canvas_result.json_data is not None:

                        rectangles = [
                            obj for obj in
                            canvas_result.json_data.get("objects", [])
                            if obj.get("type") == "rect"
                        ]

                        if rectangles:

                            rect = rectangles[-1]

                            scale_x = (
                                original_image.width / canvas_width
                            )

                            scale_y = (
                                original_image.height / canvas_height
                            )

                            left = int(rect["left"] * scale_x)
                            top = int(rect["top"] * scale_y)

                            right = int(
                                (
                                    rect["left"] + rect["width"]
                                ) * scale_x
                            )

                            bottom = int(
                                (
                                    rect["top"] + rect["height"]
                                ) * scale_y
                            )

                            left = max(0, left)
                            top = max(0, top)
                            right = min(
                                original_image.width,
                                right
                            )
                            bottom = min(
                                original_image.height,
                                bottom
                            )

                            if right > left and bottom > top:

                                selected_image = original_image.crop(
                                    (left, top, right, bottom)
                                )

                                region_selected = True

                                st.success(
                                    "Đã chọn vùng dữ liệu."
                                )

                                st.image(
                                    selected_image,
                                    caption="Vùng thông tin được chọn",
                                    use_container_width=True
                                )

                st.markdown(
                    "#### 4. Thiết lập trường cần trích xuất"
                )

                field_name = st.text_input(
                    "Tên trường dữ liệu cần trích xuất",
                    placeholder=(
                        "Ví dụ: Vessel Name, IMO Number, "
                        "Crew Name, Date of Arrival..."
                    ),
                    key=f"field_{ship['id']}"
                )

                extraction_type_label = st.radio(
                    "Loại dữ liệu",
                    ["Một giá trị", "Bảng nhiều dòng"],
                    horizontal=True,
                    key=f"type_{ship['id']}"
                )

                extraction_type = (
                    "table"
                    if extraction_type_label == "Bảng nhiều dòng"
                    else "single"
                )

                st.caption(
                    "Bước này chỉ trích xuất dữ liệu. "
                    "Việc chọn Sheet và vị trí ghi Excel được "
                    "thực hiện ở Bước 4."
                )

                st.divider()

                
                if st.button(
                    "🤖 Bắt đầu trích xuất bằng AI",
                    type="primary",
                    use_container_width=True,
                    key=f"run_ai_{ship['id']}"
                ):

                    if not field_name.strip():

                        st.warning(
                            "Vui lòng nhập tên trường dữ liệu."
                        )

                    elif (
                        extraction_mode
                        == "Khoanh vùng thông tin cần đọc"
                        and not region_selected
                    ):

                        st.warning(
                            "Bạn chưa khoanh vùng thông tin trên chứng từ."
                        )

                    else:

                        try:

                            with st.spinner(
                                "AI đang đọc chứng từ..."
                            ):

                                result = sgcl_extract_with_ai(
                                    selected_image,
                                    field_name.strip(),
                                    extraction_type
                                )

                            extraction_record = {
                                "id": str(uuid.uuid4()),
                                "document": selected_doc["name"],
                                "page": page_number + 1,
                                "field": field_name.strip(),
                                "type": extraction_type,
                                "result": result,
                                "sheet": NSW_SHEETS[0],
                                "cell": "A1"
                            }

                            ship["field_extractions"].append(
                                extraction_record
                            )

                            ship["review_confirmed"] = False
                            ship["excel_output_bytes"] = None

                            ship["step_status"]["3"] = "done"
                            ship["step_status"]["4"] = "pending"
                            ship["step_status"]["5"] = "pending"

                            ship["status"] = "Đã trích xuất"

                            st.success(
                                f"Đã trích xuất thành công trường: "
                                f"{field_name.strip()}"
                            )

                            st.json(result)

                        except Exception as e:

                            import traceback

                            st.error(
                                f"Lỗi khi trích xuất: {str(e)}"
                            )

                            with st.expander(
                                "Chi tiết lỗi kỹ thuật"
                            ):

                                st.code(
                                    traceback.format_exc(),
                                    language="python"
                                )

            except Exception as e:

                st.error(
                    f"Không thể xử lý chứng từ: {str(e)}"
                )

        # LỊCH SỬ TRÍCH XUẤT

        if ship["field_extractions"]:

            st.divider()
            st.markdown(
                "#### Lịch sử dữ liệu đã trích xuất"
            )

            for index, item in enumerate(
                ship["field_extractions"],
                start=1
            ):

                with st.expander(
                    f"{index}. {item['field']} – {item['document']}"
                ):

                    st.write(f"Trang: {item['page']}")
                    st.write(f"Loại dữ liệu: {item['type']}")
                    st.json(item["result"])

        # ĐIỀU HƯỚNG

        st.divider()

        c1, c2 = st.columns(2)

        with c1:

            if st.button(
                "← Quay lại",
                use_container_width=True,
                key=f"back_step3_{ship['id']}"
            ):

                ship["current_step"] = 2
                st.rerun()

        with c2:

            if st.button(
                "Xác nhận & Tiếp tục →",
                type="primary",
                use_container_width=True,
                disabled=not ship["field_extractions"],
                key=f"next_step3_{ship['id']}"
            ):

                ship["step_status"]["3"] = "done"
                ship["current_step"] = 4
                st.rerun()

    # =====================================================
    # BƯỚC 4 – KIỂM TRA VÀ SẮP XẾP DỮ LIỆU
    # =====================================================

    elif step == 4:

        st.subheader("4. Kiểm tra dữ liệu")

        st.caption(
            "Kiểm tra kết quả AI, chỉnh sửa thông tin và "
            "chọn Sheet cùng vị trí ghi dữ liệu trong biểu mẫu NSW."
        )

        records = ship.get("field_extractions", [])

        if not records:

            st.warning(
                "Chưa có dữ liệu trích xuất. "
                "Vui lòng thực hiện Bước 3 trước."
            )

        else:

            import pandas as pd

            st.markdown(
                "#### Danh sách dữ liệu đã trích xuất"
            )

            st.info(
                "Mỗi trường dữ liệu có thể được đưa vào một Sheet "
                "và một vị trí Excel khác nhau."
            )

            all_valid = True

            for index, record in enumerate(records):

                record.setdefault("id", str(uuid.uuid4()))
                record.setdefault("sheet", NSW_SHEETS[0])
                record.setdefault("cell", "A1")

                with st.container(border=True):

                    st.markdown(
                        f"**{index + 1}. {record['field']}**"
                    )

                    st.caption(
                        f"Chứng từ: {record['document']} "
                        f"· Trang: {record['page']}"
                    )

                    current_sheet = record.get(
                        "sheet",
                        NSW_SHEETS[0]
                    )

                    if current_sheet not in NSW_SHEETS:
                        current_sheet = NSW_SHEETS[0]

                    selected_sheet = st.selectbox(
                        "Chọn Sheet NSW",
                        NSW_SHEETS,
                        index=NSW_SHEETS.index(current_sheet),
                        key=f"review_sheet_{ship['id']}_{record['id']}"
                    )

                    old_cell = record.get("cell", "A1")

                    start_cell = st.text_input(
                        "Ô bắt đầu ghi dữ liệu",
                        value=old_cell,
                        placeholder="Ví dụ: B5",
                        key=f"review_cell_{ship['id']}_{record['id']}"
                    )

                    if selected_sheet != record.get("sheet"):
                        record["sheet"] = selected_sheet
                        ship["review_confirmed"] = False

                    if start_cell.strip().upper() != str(
                        record.get("cell", "A1")
                    ).strip().upper():

                        record["cell"] = start_cell.strip().upper()
                        ship["review_confirmed"] = False

                    # KIỂM TRA ĐỊA CHỈ Ô

                    try:

                        openpyxl.utils.cell.coordinate_from_string(
                            start_cell.strip().upper()
                        )

                    except Exception:

                        st.error(
                            "Địa chỉ ô Excel không hợp lệ. "
                            "Ví dụ hợp lệ: A1, B5, C10."
                        )

                        all_valid = False

                    # DỮ LIỆU ĐƠN

                    if record["type"] == "single":

                        old_value = str(
                            record["result"].get("value", "")
                        )

                        edited_value = st.text_area(
                            "Giá trị trích xuất – có thể chỉnh sửa",
                            value=old_value,
                            height=100,
                            key=f"review_value_{ship['id']}_{record['id']}"
                        )

                        if edited_value != old_value:

                            record["result"]["value"] = edited_value
                            ship["review_confirmed"] = False

                        if not edited_value.strip():

                            st.warning(
                                "Giá trị đang bị trống."
                            )

                            all_valid = False

                    # DỮ LIỆU BẢNG

                    else:

                        result = record.get("result", {})

                        headers = result.get("headers", [])
                        rows = result.get("rows", [])

                        width = max(
                            len(headers),
                            max(
                                (len(row) for row in rows),
                                default=0
                            ),
                            1
                        )

                        columns = []

                        for col_index in range(width):

                            if col_index < len(headers):
                                column_name = str(
                                    headers[col_index]
                                ).strip()
                            else:
                                column_name = ""

                            if not column_name:
                                column_name = (
                                    f"Cột {col_index + 1}"
                                )

                            if column_name in columns:
                                column_name = (
                                    f"{column_name}_{col_index + 1}"
                                )

                            columns.append(column_name)

                        normalized_rows = [
                            [
                                str(value)
                                if value is not None
                                else ""
                                for value in row
                            ] + [""] * (width - len(row))
                            for row in rows
                        ]

                        df = pd.DataFrame(
                            normalized_rows,
                            columns=columns
                        )

                        st.caption(
                            "Có thể sửa trực tiếp từng ô hoặc thêm dòng mới."
                        )

                        edited_df = st.data_editor(
                            df,
                            use_container_width=True,
                            hide_index=True,
                            num_rows="dynamic",
                            key=f"review_table_{ship['id']}_{record['id']}"
                        )

                        edited_rows = [
                            [
                                str(value)
                                if value is not None
                                else ""
                                for value in row
                            ]
                            for row in edited_df.fillna("").values.tolist()
                        ]

                        new_result = {
                            "headers": columns,
                            "rows": edited_rows
                        }

                        if new_result != record["result"]:

                            record["result"] = new_result
                            ship["review_confirmed"] = False

                        if not any(
                            any(
                                str(value).strip()
                                for value in row
                            )
                            for row in edited_rows
                        ):

                            st.warning(
                                "Bảng chưa có dữ liệu để xuất."
                            )

                            all_valid = False

                    st.caption(
                        f"Đích ghi dữ liệu: {selected_sheet} – "
                        f"{start_cell.strip().upper()}"
                    )

            # XÁC NHẬN DỮ LIỆU

            st.divider()

            st.markdown("#### Xác nhận dữ liệu")

            if all_valid:

                st.success(
                    "Các trường dữ liệu đã có thông tin "
                    "và địa chỉ ô Excel hợp lệ."
                )

            else:

                st.warning(
                    "Vui lòng kiểm tra các trường bị trống "
                    "hoặc địa chỉ ô Excel chưa hợp lệ."
                )

            if st.button(
                "Xác nhận dữ liệu & Tiếp tục →",
                type="primary",
                use_container_width=True,
                disabled=not all_valid,
                key=f"confirm_step4_{ship['id']}"
            ):

                ship["review_confirmed"] = True
                ship["step_status"]["4"] = "done"
                ship["status"] = "Đã xác nhận"
                ship["current_step"] = 5
                st.rerun()

        # ĐIỀU HƯỚNG

        st.divider()

        if st.button(
            "← Quay lại Bước 3",
            use_container_width=True,
            key=f"back_step4_{ship['id']}"
        ):

            ship["current_step"] = 3
            st.rerun()

    # =====================================================
    # BƯỚC 5 – XUẤT FILE EXCEL NSW
    # =====================================================

    elif step == 5:

        st.subheader("5. Xuất file Excel NSW")

        st.caption(
            "Tải biểu mẫu Excel NSW gốc, hệ thống sẽ tự động "
            "điền dữ liệu đã xác nhận vào các Sheet tương ứng."
        )

        if not ship.get("review_confirmed"):

            st.warning(
                "Bạn cần hoàn tất kiểm tra và xác nhận dữ liệu "
                "ở Bước 4 trước khi xuất file."
            )

        else:

            st.markdown(
                "#### 1. Kiểm tra dữ liệu sẽ xuất"
            )

            records = ship.get("field_extractions", [])

            if not records:

                st.warning("Không có dữ liệu để xuất.")

            else:

                summary_data = []

                for record in records:

                    summary_data.append({
                        "Trường dữ liệu": record["field"],
                        "Sheet": record.get("sheet", ""),
                        "Ô bắt đầu": record.get("cell", ""),
                        "Loại dữ liệu": (
                            "Bảng nhiều dòng"
                            if record["type"] == "table"
                            else "Một giá trị"
                        )
                    })

                st.dataframe(
                    summary_data,
                    use_container_width=True,
                    hide_index=True
                )

                st.divider()

                # TẢI FILE EXCEL GỐC

                st.markdown(
                    "#### 2. Tải lên file Excel NSW gốc"
                )

                st.info(
                    "Khu vực này được chừa sẵn để bà tải file Excel NSW "
                    "gốc lên khi đã có mẫu. Hệ thống chưa yêu cầu "
                    "một file mẫu cố định."
                )

                uploaded_template = st.file_uploader(
                    "Chọn file Excel NSW gốc (.xlsx)",
                    type=["xlsx"],
                    key=f"original_template_{ship['id']}"
                )

                if uploaded_template is not None:

                    template_bytes = uploaded_template.getvalue()

                    template_hash = hashlib.sha256(
                        template_bytes
                    ).hexdigest()

                    if template_hash != ship.get(
                        "excel_template_hash"
                    ):

                        ship["excel_template_bytes"] = template_bytes
                        ship["excel_template_hash"] = template_hash

                        ship["excel_template_name"] = (
                            uploaded_template.name
                        )

                        ship["excel_output_bytes"] = None

                        st.success(
                            "Đã tiếp nhận file Excel NSW gốc."
                        )

                # TẠO FILE EXCEL

                if ship.get("excel_template_bytes"):

                    st.markdown(
                        "#### 3. Tạo file Excel hoàn chỉnh"
                    )

                    if st.button(
                        "📊 Tạo file Excel NSW",
                        type="primary",
                        use_container_width=True,
                        key=f"generate_excel_{ship['id']}"
                    ):

                        try:

                            template_bytes = ship[
                                "excel_template_bytes"
                            ]

                            workbook = openpyxl.load_workbook(
                                BytesIO(template_bytes),
                                read_only=True
                            )

                            available_sheets = workbook.sheetnames
                            workbook.close()

                            missing_sheets = sorted({
                                record.get("sheet", "")
                                for record in records
                                if record.get("sheet", "")
                                not in available_sheets
                            })

                            if missing_sheets:

                                st.error(
                                    "File Excel gốc không có các Sheet "
                                    "đã chọn ở Bước 4: "
                                    + ", ".join(missing_sheets)
                                )

                            else:

                                output_bytes = template_bytes

                                with st.spinner(
                                    "Đang ghi dữ liệu vào biểu mẫu NSW..."
                                ):

                                    for record in records:

                                        output_bytes = sgcl_write_to_excel(
                                            output_bytes,
                                            record["sheet"],
                                            record["cell"],
                                            record["result"],
                                            record["type"]
                                        )

                                ship["excel_output_bytes"] = output_bytes

                                safe_voyage = "".join(
                                    char
                                    for char in str(ship["voyage_no"])
                                    if char.isalnum() or char in "-_"
                                )

                                ship["excel_output_name"] = (
                                    f"SGCL_NSW_{ship['imo']}_"
                                    f"{safe_voyage}.xlsx"
                                )

                                ship["step_status"]["5"] = "done"
                                ship["status"] = "Hoàn thành"

                                st.success(
                                    "Đã tạo file Excel NSW thành công!"
                                )

                        except Exception as e:

                            st.error(
                                f"Không thể tạo file Excel: {str(e)}"
                            )

                else:

                    st.warning(
                        "Vui lòng tải file Excel NSW gốc lên "
                        "để hệ thống có thể tạo file kết quả."
                    )

                # TẢI FILE KẾT QUẢ

                if ship.get("excel_output_bytes"):

                    st.divider()

                    st.markdown(
                        "#### 4. Tải file Excel hoàn chỉnh"
                    )

                    st.download_button(
                        "⬇️ Tải xuống file Excel NSW",
                        data=ship["excel_output_bytes"],
                        file_name=ship.get(
                            "excel_output_name",
                            "SGCL_NSW.xlsx"
                        ),
                        mime=(
                            "application/vnd.openxmlformats-officedocument."
                            "spreadsheetml.sheet"
                        ),
                        type="primary",
                        use_container_width=True,
                        key=f"download_final_excel_{ship['id']}"
                    )

        # ĐIỀU HƯỚNG

        st.divider()

        if st.button(
            "← Quay lại Bước 4",
            use_container_width=True,
            key=f"back_step5_{ship['id']}"
        ):

            ship["current_step"] = 4
            st.rerun()
