import streamlit as st
import sqlite3
import base64
import uuid
from datetime import datetime
import time
import hashlib
from pathlib import Path
import os
from io import BytesIO
from PIL import Image
import fitz
import openpyxl
import sys
import re
import json
import pandas as pd

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
# SQLITE DATABASE - SGCL
# =========================================================

DB_DIR = Path(__file__).resolve().parent / "data"
DB_DIR.mkdir(parents=True, exist_ok=True)

DB_PATH = DB_DIR / "sgcl.db"


def get_db_connection():
    """Tạo kết nối SQLite."""
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    return conn


def init_database():
    """Khởi tạo bảng lưu hồ sơ tàu."""

    with get_db_connection() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS ship_records (
                ship_id TEXT PRIMARY KEY,
                imo TEXT NOT NULL UNIQUE,
                created_at REAL,
                updated_at TEXT NOT NULL,
                data_json TEXT NOT NULL
            )
        """)

        conn.commit()


def _json_default(value):
    """Chuyển bytes thành base64 để lưu file Excel trong SQLite."""
    if isinstance(value, (bytes, bytearray)):
        return {"__sgcl_bytes__": base64.b64encode(bytes(value)).decode("ascii")}
    raise TypeError(f"Không thể lưu kiểu dữ liệu: {type(value).__name__}")


def _json_object_hook(value):
    """Khôi phục bytes khi đọc hồ sơ từ SQLite."""
    if set(value.keys()) == {"__sgcl_bytes__"}:
        return base64.b64decode(value["__sgcl_bytes__"])
    return value


def load_ships_from_db():
    """Đọc toàn bộ hồ sơ tàu từ SQLite."""

    with get_db_connection() as conn:
        rows = conn.execute("""
            SELECT data_json
            FROM ship_records
            ORDER BY created_at DESC
        """).fetchall()

    ships = []

    for row in rows:
        try:
            ships.append(json.loads(row["data_json"], object_hook=_json_object_hook))
        except (json.JSONDecodeError, TypeError):
            continue

    return ships


def save_ship_to_db(ship):
    """Lưu hoặc cập nhật một hồ sơ tàu."""

    now = datetime.now().isoformat(timespec="seconds")

    data_json = json.dumps(
        ship,
        ensure_ascii=False,
        default=_json_default
    )

    with get_db_connection() as conn:
        conn.execute("""
            INSERT INTO ship_records (
                ship_id,
                imo,
                created_at,
                updated_at,
                data_json
            )
            VALUES (?, ?, ?, ?, ?)

            ON CONFLICT(ship_id) DO UPDATE SET
                imo = excluded.imo,
                created_at = excluded.created_at,
                updated_at = excluded.updated_at,
                data_json = excluded.data_json
        """, (
            ship["id"],
            ship["imo"],
            ship.get("created_at", datetime.now().timestamp()),
            now,
            data_json
        ))

        conn.commit()


def delete_ship_from_db(ship_id):
    """Xóa hồ sơ tàu khỏi SQLite."""

    with get_db_connection() as conn:
        conn.execute(
            "DELETE FROM ship_records WHERE ship_id = ?",
            (ship_id,)
        )
        conn.commit()


def save_all_ships_to_db():
    """
    Đồng bộ hồ sơ xuống SQLite.
    Khi đang ở trang chi tiết chỉ lưu hồ sơ đang mở để tránh ghi lại
    toàn bộ dữ liệu (kể cả file Excel) của mọi tàu sau mỗi thao tác.
    """
    ships = st.session_state.get("ships", [])
    current_id = st.session_state.get("current_ship_id")

    if st.session_state.get("page") == "detail" and current_id:
        ships = [s for s in ships if s.get("id") == current_id]

    for item in ships:
        save_ship_to_db(item)


# =========================================================
# THÔNG BÁO GIỮ LẠI SAU st.rerun()
# =========================================================

def flash(message, kind="success"):
    """Lưu thông báo để hiển thị sau khi trang được tải lại."""
    st.session_state.setdefault("_flash", []).append((kind, message))


def show_flash():
    for kind, message in st.session_state.pop("_flash", []):
        getattr(st, kind, st.info)(message)


# Khởi tạo cơ sở dữ liệu khi ứng dụng chạy
init_database()
# =========================================================
# 1. CẤU HÌNH LƯU TRỮ
# =========================================================

APP_DIR = Path(__file__).resolve().parent
BASE_UPLOAD_DIR = APP_DIR / "uploads"
BASE_UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
ALLOWED_EXTENSIONS = {".pdf"}
MAX_FILE_SIZE = 20 * 1024 * 1024


def get_ship_folder(ship):
    folder = BASE_UPLOAD_DIR / f"{ship['imo']}_{ship['id']}"
    folder.mkdir(parents=True, exist_ok=True)
    return folder


def resolve_doc_path(doc):
    """
    Trả về đường dẫn tuyệt đối của chứng từ.
    Hồ sơ cũ lưu đường dẫn tương đối 'uploads/...' nên cần quy về thư mục app.
    """
    path = Path(doc.get("path", ""))
    if not path.is_absolute():
        path = APP_DIR / path
    return path


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

def sgcl_safe_zoom(page, zoom, max_side=2200):
    """
    Giới hạn kích thước ảnh render. Bản scan thường có trang rất lớn
    (VD 2480x3508 điểm); phóng to thêm sẽ tạo ảnh hàng trăm MB và làm
    app hết RAM rồi bị tắt.
    """
    rect = page.rect
    longest = max(rect.width, rect.height, 1)
    return max(0.2, min(zoom, max_side / longest))


@st.cache_data(show_spinner=False)
def render_pdf_pages(file_bytes, zoom=1.6):
    """Chuyển các trang PDF thành ảnh PNG."""
    images = []

    with fitz.open(stream=file_bytes, filetype="pdf") as pdf:
        for page in pdf:
            z = sgcl_safe_zoom(page, zoom, max_side=1800)
            pix = page.get_pixmap(matrix=fitz.Matrix(z, z))
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

        else:
            st.warning("Chỉ hỗ trợ xem trước file PDF.")

    except Exception as e:
        st.error(f"Không thể xem trước chứng từ: {str(e)}")


# =========================================================
# 3. CÁC HÀM HỖ TRỢ PDF/OCR - BƯỚC 3
# =========================================================

def sgcl_get_document_bytes(doc):
    """Đọc chứng từ đã tải lên ở Bước 2."""
    file_path = resolve_doc_path(doc)

    if not file_path.exists():
        raise FileNotFoundError(
            f"Không tìm thấy chứng từ: {doc['name']}"
        )

    return file_path.read_bytes()


def sgcl_render_pages(file_bytes, file_name):
    """Chuyển từng trang PDF thành ảnh nội bộ để xem trước và OCR."""
    if Path(file_name).suffix.lower() != ".pdf":
        raise ValueError("Chỉ hỗ trợ chứng từ định dạng PDF.")

    images = []
    with fitz.open(stream=file_bytes, filetype="pdf") as pdf:
        for page in pdf:
            z = sgcl_safe_zoom(page, 2, max_side=2200)
            pix = page.get_pixmap(matrix=fitz.Matrix(z, z), alpha=False)
            images.append(Image.open(BytesIO(pix.tobytes("png"))).convert("RGB"))
    return images


def sgcl_extract_with_ocr(
    image,
    field_name,
    extraction_type="single"
):
    """
    Nhận diện chữ trong trang PDF bằng Tesseract OCR.
    """

    import pytesseract
    import re
    from pytesseract import Output

    # Chuyển ảnh về RGB để OCR
    if image.mode != "RGB":
        image = image.convert("RGB")

    # Phóng to ảnh để cải thiện nhận diện
    scale = 2
    image = image.resize(
        (image.width * scale, image.height * scale)
    )

    # OCR toàn bộ nội dung
    ocr_data = pytesseract.image_to_data(
        image,
        lang=sgcl_tesseract_languages(),
        config="--oem 3 --psm 6",
        output_type=Output.DICT
    )

    # Gom các từ thành từng dòng
    lines = {}

    for i, word in enumerate(ocr_data["text"]):
        word = word.strip()

        if not word:
            continue

        key = (
            ocr_data["block_num"][i],
            ocr_data["par_num"][i],
            ocr_data["line_num"][i]
        )

        if key not in lines:
            lines[key] = []

        lines[key].append({
            "text": word,
            "left": ocr_data["left"][i],
            "width": ocr_data["width"][i]
        })

    # Sắp xếp từ theo vị trí trái sang phải
    ocr_lines = []

    for words in lines.values():
        words.sort(key=lambda x: x["left"])

        line_text = " ".join(
            word["text"] for word in words
        )

        ocr_lines.append({
            "text": line_text,
            "words": words
        })

    full_text = "\n".join(
        line["text"] for line in ocr_lines
    )

    # ==================================
    # TRƯỜNG HỢP TRÍCH XUẤT BẢNG
    # ==================================
    if extraction_type == "table":

        rows = []

        for line in ocr_lines:
            words = line["words"]

            if not words:
                continue

            cells = []
            current_cell = words[0]["text"]
            previous_right = (
                words[0]["left"] + words[0]["width"]
            )

            for word in words[1:]:

                gap = word["left"] - previous_right

                # Khoảng cách lớn được xem là ranh giới cột
                if gap > 35:
                    cells.append(current_cell)
                    current_cell = word["text"]
                else:
                    current_cell += " " + word["text"]

                previous_right = (
                    word["left"] + word["width"]
                )

            cells.append(current_cell)
            rows.append(cells)

        if not rows:
            return {
                "headers": [],
                "rows": []
            }

        headers = rows[0]
        data_rows = rows[1:]

        return {
            "headers": headers,
            "rows": data_rows
        }

    # ==================================
    # TRƯỜNG HỢP TRÍCH XUẤT THÔNG TIN
    # ==================================

    field_normalized = re.sub(
        r"[^a-z0-9]",
        "",
        field_name.lower()
    )

    extracted_value = ""

    # Tìm dòng có chứa tên trường cần lấy
    for index, line in enumerate(ocr_lines):

        line_text = line["text"]

        line_normalized = re.sub(
            r"[^a-z0-9]",
            "",
            line_text.lower()
        )

        if field_normalized and field_normalized in line_normalized:

            # Lấy nội dung phía sau dấu :
            if ":" in line_text:
                extracted_value = line_text.split(
                    ":", 1
                )[1].strip()

            # Nếu không có giá trị trên cùng dòng,
            # thử lấy dòng kế tiếp
            if not extracted_value and index + 1 < len(ocr_lines):
                extracted_value = ocr_lines[
                    index + 1
                ]["text"]

            break

    return {
        "value": extracted_value
    }

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
                "Không trích xuất được dữ liệu bảng từ PDF."
            )

        start = openpyxl.utils.cell.coordinate_from_string(
            start_cell.upper()
        )

        start_col = openpyxl.utils.column_index_from_string(
            start[0]
        )

        start_row = start[1]

        for col_index, header in enumerate(headers):
            sgcl_write_cell(ws, start_row, start_col + col_index, header)

        for row_index, row_data in enumerate(rows, start=1):
            for col_index, value in enumerate(row_data):
                sgcl_write_cell(
                    ws, start_row + row_index, start_col + col_index, value
                )

    else:
        col_letter, row_number = openpyxl.utils.cell.coordinate_from_string(
            start_cell.upper()
        )
        sgcl_write_cell(
            ws,
            row_number,
            openpyxl.utils.column_index_from_string(col_letter),
            result.get("value", "")
        )

    output = BytesIO()
    workbook.save(output)
    workbook.close()

    return output.getvalue()


# =========================================================
# 4. DANH SÁCH SHEET NSW
# =========================================================

# Chỉ 10 sheet nghiệp vụ được phép ánh xạ/ghi dữ liệu.
# Tên được giữ đúng như trong file TemplateNSW.xlsx.
NSW_SHEETS = [
    "Notice ship to port",
    "ConfirmationOfArrival",
    "General Declaration",
    "Goods Declaration",
    "Crew list",
    "Crew's Effects Declaration",
    "Ship's store declaration",
    "Declaration For Ship Security",
    "Declaration For Plant",
    "Declaration For Health"
]

# ---------------------------------------------------------
# VỊ TRÍ CÁC Ô DÙNG CHUNG TRONG TemplateNSW.xlsx
# (đối chiếu trực tiếp với file mẫu: B5 = chức năng chứng từ,
#  3 ô Nơi ký / Ngày ký / Người ký ở cuối phần thông tin chung)
# ---------------------------------------------------------
NSW_COMMON_CELLS = {
    "Notice ship to port":           {"place": "B30", "date": "B31", "signer": "B32", "date_fmt": "%d/%m/%Y %H:%M"},
    "ConfirmationOfArrival":         {"place": "B8",  "date": "B9",  "signer": "B10", "date_fmt": "%d/%m/%Y %H:%M"},
    "General Declaration":           {"place": "B32", "date": "B33", "signer": "B34", "date_fmt": "%d/%m/%Y %H:%M"},
    "Goods Declaration":             {},
    "Crew list":                     {"place": "B6",  "date": "B7",  "signer": "B8",  "date_fmt": "%d/%m/%Y %H:%M"},
    "Crew's Effects Declaration":    {"place": "B6",  "date": "B7",  "signer": "B8",  "date_fmt": "%d/%m/%Y %H:%M"},
    "Ship's store declaration":      {"place": "B8",  "date": "B9",  "signer": "B10", "date_fmt": "%d/%m/%Y %H:%M"},
    "Declaration For Ship Security": {"place": "B20", "date": "B21", "signer": "B22", "date_fmt": "%d/%m/%Y %H:%M"},
    "Declaration For Plant":         {"place": "B13", "date": "B14", "signer": "B15", "date_fmt": "%d/%m/%Y"},
    "Declaration For Health":        {"place": "B20", "date": "B21", "signer": "B22", "date_fmt": "%d/%m/%Y"},
}

NSW_FUNCTION_CODES = {
    "CN01": "Thêm mới bản khai (Create)",
    "CN02": "Sửa bản khai (Update)",
    "CN03": "Xóa bản khai (Delete)",
}

NSW_SIGNERS = [
    "Thuyền trưởng/Master",
    "Đại lý/ShipAgency",
    "Chủ tàu/ShipOwner",
]

# ---------------------------------------------------------
# ÁNH XẠ MẶC ĐỊNH: (loại chứng từ, trường) -> vị trí trong template.
# "extra" = các ô ghi thêm (cùng dữ liệu xuất hiện ở nhiều bản khai).
# Ưu tiên: vị trí đã dùng ở hồ sơ trước > mặc định dưới đây.
# Trường bảng có tên "Bảng dữ liệu - <loại chứng từ>" như ở Bước 3.
# ---------------------------------------------------------
_GD, _NT, _CA = "General Declaration", "Notice ship to port", "ConfirmationOfArrival"

DEFAULT_NSW_MAPPING_LIST = [
    # Ship's Particulars
    ("Ship’s Particulars", "Air draft", _GD, "B12", [(_NT, "B12")], None),

    # Notice of Arrival (B6 = các cảng trước, phân cách bởi ';')
    ("Notice of Arrival", "Last port of call", _GD, "B6", [(_NT, "B6")], None),

    # Arrival Report
    ("Arrival Report", "Time of arrival", _GD, "B24", [(_NT, "B24"), (_CA, "B6")], None),
    ("Arrival Report", "All fast", _GD, "B25", [(_NT, "B25")], None),
    ("Arrival Report", "P.O.B", _GD, "B26", [(_NT, "B26")], None),
    ("Arrival Report", "M.D.O", _GD, "B29", [(_NT, "B27")], None),
    ("Arrival Report", "F.O", _GD, "B30", [(_NT, "B28")], None),
    ("Arrival Report", "FreshWater", _GD, "B31", [(_NT, "B29")], None),

    # Kế hoạch cập cầu / Last Port Clearance
    ("Kế hoạch cập cầu", "Berthing Position", _GD, "B21", [(_NT, "B21")], None),
    ("Last Port Clearance", "No.", _GD, "B23", [(_NT, "B23")], None),

    # Bill of Lading -> dòng hàng đầu tiên của Goods Declaration
    ("Bill of Lading", "B/L No.", "Goods Declaration", "A10", [], None),
    ("Bill of Lading", "Shipper", "Goods Declaration", "B10", [], None),
    ("Bill of Lading", "Consignee", "Goods Declaration", "C10", [], None),
    ("Bill of Lading", "Notify Party", "Goods Declaration", "D10", [], None),
    ("Bill of Lading", "Shipper’s Description of Goods", "Goods Declaration", "I10", [], None),
    ("Bill of Lading", "Gross weight", "Goods Declaration", "K10", [], None),
    ("Bill of Lading", "Port of Discharge", "Goods Declaration", "P10", [], None),
    ("Bill of Lading", "Port of Loading", "Goods Declaration", "R10", [], None),

    # ISSC -> Declaration For Ship Security
    ("International Ship Security Certificate",
     "Date of initial / renewal verification on which this certificate is based",
     "Declaration For Ship Security", "B10", [], None),
    ("International Ship Security Certificate", "This certificate is valid until",
     "Declaration For Ship Security", "B11", [], None),

    # Ship's Stores List
    ("Ship’s Stores List", "Number of persons on board", "Ship's store declaration", "B6", [], None),
    ("Ship’s Stores List", "Period of stay", "Ship's store declaration", "B7", [], None),
    ("Ship’s Stores List", "Bảng dữ liệu - Ship’s Stores List", "Ship's store declaration", "A13", [], {
        "Name of article": "A", "Quantity": "B", "Unit": "C", "Place of storage": "D"
    }),

    # Maritime Declaration of Health
    ("Maritime Declaration of Health", "Arriving from", "Declaration For Health", "B8", [], None),
    ("Maritime Declaration of Health", "Sailing to", "Declaration For Health", "B9", [], None),
    ("Maritime Declaration of Health", "Valid SSCEC/SSCC carried on board?", "Declaration For Health", "B10", [], None),
    ("Maritime Declaration of Health", "Issued at", "Declaration For Health", "B11", [], None),
    ("Maritime Declaration of Health", "Re-inspection required?", "Declaration For Health", "B13", [], None),

    # Crew List (template tách Họ/Tên, Ngày/Nơi sinh, Số/Hạn hộ chiếu)
    ("Crew List", "Bảng dữ liệu - Crew List", "Crew list", "A11", [], {
        "Family name, given name": "A",
        "Rank or rating": "C",
        "Nationality": "D",
        "Date and place of birth (dd/mm/yyyy)": "E",
        "Passport No and Passport expire date": "H",
    }),

    # Crew's Effects
    ("Crew’s Effects List", "Bảng dữ liệu - Crew’s Effects List", "Crew's Effects Declaration", "A11", [], {
        "Family name, given names": "A",
        "Rank or rating": "C",
        "Effects ineligible for relief from customs duties and taxes or subject to prohibitions or restrictions": "D",
    }),

    # Voyage Memory -> bảng cảng đã ghé của Declaration For Ship Security
    ("Voyage Memory", "Bảng dữ liệu - Voyage Memory", "Declaration For Ship Security", "A25", [], {
        "Port": "B", "Date of Arrival": "C", "Date of Departure": "D"
    }),
]

# =========================================================
# CẤU HÌNH TRƯỜNG DỮ LIỆU CHỨNG TỪ
# =========================================================

BASE_DIR = Path(__file__).resolve().parent
CONFIG_PATH = BASE_DIR / "sgcl_field_config.json"
TEMPLATE_PATH = BASE_DIR / "TemplateNSW.xlsx"

DEFAULT_FIELD_CONFIG = {
    "Ship’s Particulars": {
        "aliases": ["ship particulars", "ship particular"],
        "fields": [
            {"name": "Arrival draft (F/A)", "type": "single"},
            {"name": "Air draft", "type": "single"},
        ]
    },
    "Notice of Arrival": {
        "aliases": ["notice of arrival", "noa"],
        "fields": [
            {"name": "Last port of call", "type": "single"},
            {"name": "Port of arrival", "type": "single"},
            {"name": "Departure from last port", "type": "single"},
        ]
    },
    "Arrival Report": {
        "aliases": ["arrival report"],
        "fields": [
            {"name": "Time of arrival", "type": "single"},
            {"name": "All fast", "type": "single"},
            {"name": "P.O.B", "type": "single"},
            {"name": "F.O", "type": "single"},
            {"name": "M.D.O", "type": "single"},
            {"name": "FreshWater", "type": "single"},
        ]
    },
    "Kế hoạch cập cầu": {
        "aliases": ["berthing plan", "berthing position"],
        "fields": [
            {"name": "Berthing Position", "type": "single"},
        ]
    },
    "Cargo Manifest": {
        "aliases": ["cargo manifest", "manifest"],
        "fields": [
            {"name": "Port of discharge", "type": "single"},
        ]
    },
    "Last Port Clearance": {
        "aliases": ["last port clearance"],
        "fields": [
            {"name": "No.", "type": "single"},
        ]
    },
    "Bill of Lading": {
        "aliases": ["bill of lading", "b/l"],
        "fields": [
            {"name": "B/L No.", "type": "single"},
            {"name": "Shipper", "type": "single"},
            {"name": "Consignee", "type": "single"},
            {"name": "Notify Party", "type": "single"},
            {"name": "Shipper’s Description of Goods", "type": "single"},
            {"name": "Gross weight", "type": "single"},
            {"name": "Port of Loading", "type": "single"},
            {"name": "Port of Discharge", "type": "single"},
        ]
    },
    "International Ship Security Certificate": {
        "aliases": ["international ship security", "ship security certificate"],
        "fields": [
            {"name": "Date of initial / renewal verification on which this certificate is based", "type": "single"},
            {"name": "This certificate is valid until", "type": "single"},
        ]
    },
    "Voyage Memory": {
        "aliases": ["voyage memory"],
        "fields": [
            {"name": "Port", "type": "table"},
            {"name": "Date of Arrival", "type": "table"},
            {"name": "Date of Departure", "type": "table"},
        ]
    },
    "Crew List": {
        "aliases": ["crew list", "crewlist"],
        "fields": [
            {"name": "IMO Number", "type": "single"},
            {"name": "Family name, given name", "type": "table"},
            {"name": "Rank or rating", "type": "table"},
            {"name": "Nationality", "type": "table"},
            {"name": "Date and place of birth (dd/mm/yyyy)", "type": "table"},
            {"name": "Passport No and Passport expire date", "type": "table"},
        ]
    },
    "Crew’s Effects List": {
        "aliases": ["crew effects", "crew's effects"],
        "fields": [
            {"name": "Family name, given names", "type": "table"},
            {"name": "Rank or rating", "type": "table"},
            {"name": "Effects ineligible for relief from customs duties and taxes or subject to prohibitions or restrictions", "type": "table"},
        ]
    },
    "Ship’s Stores List": {
        "aliases": ["ship stores", "ship's stores", "stores list"],
        "fields": [
            {"name": "Number of persons on board", "type": "single"},
            {"name": "Period of stay", "type": "single"},
            {"name": "Name of article", "type": "table"},
            {"name": "Quantity", "type": "table"},
            {"name": "Unit", "type": "table"},
            {"name": "Place of storage", "type": "table"},
        ]
    },
    "Maritime Declaration of Health": {
        "aliases": ["maritime declaration of health", "declaration of health"],
        "fields": [
            {"name": "Arriving from", "type": "single"},
            {"name": "Sailing to", "type": "single"},
            {"name": "Valid SSCEC/SSCC carried on board?", "type": "single"},
            {"name": "Issued at", "type": "single"},
            {"name": "Re-inspection required?", "type": "single"},
            {"name": "List of ports of call from commencement of voyage or within past thirty days", "type": "table"},
        ]
    },
}


def sgcl_load_field_config():
    """Giữ cấu hình người dùng, đồng thời bổ sung trường mới từ mặc định."""
    default_copy = json.loads(json.dumps(DEFAULT_FIELD_CONFIG, ensure_ascii=False))
    if CONFIG_PATH.exists():
        try:
            saved = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
            if not isinstance(saved, dict):
                saved = {}
            for doc_type, default_item in default_copy.items():
                if doc_type not in saved or not isinstance(saved[doc_type], dict):
                    saved[doc_type] = default_item
                    continue
                saved_fields = saved[doc_type].setdefault("fields", [])
                existing_names = {sgcl_normalize(x.get("name", "")) for x in saved_fields if isinstance(x, dict)}
                for field in default_item.get("fields", []):
                    if sgcl_normalize(field.get("name", "")) not in existing_names:
                        saved_fields.append(field)
                if not saved[doc_type].get("aliases"):
                    saved[doc_type]["aliases"] = default_item.get("aliases", [])
            CONFIG_PATH.write_text(json.dumps(saved, ensure_ascii=False, indent=2), encoding="utf-8")
            return saved
        except Exception:
            pass
    CONFIG_PATH.write_text(json.dumps(default_copy, ensure_ascii=False, indent=2), encoding="utf-8")
    return default_copy


def sgcl_save_field_config(config):
    CONFIG_PATH.write_text(
        json.dumps(config, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )


def sgcl_normalize(text):
    import unicodedata
    text = unicodedata.normalize("NFKD", str(text or ""))
    text = "".join(c for c in text if not unicodedata.combining(c))
    text = text.lower().replace("’", "'")
    return re.sub(r"[^a-z0-9]+", "", text)


def sgcl_get_document_type(filename, config):
    filename = sgcl_normalize(Path(filename).stem)

    matches = []
    for doc_type, item in config.items():
        for alias in item.get("aliases", []):
            alias_key = sgcl_normalize(alias)
            if alias_key and alias_key in filename:
                matches.append((len(alias_key), doc_type))

    if matches:
        return sorted(matches, reverse=True)[0][1]

    return None


def sgcl_words_to_rows(words, y_tolerance=3.0):
    """
    Ghép các từ của text layer thành từng dòng theo tọa độ thực tế.

    get_text("text") của PyMuPDF thường tách nhãn và giá trị nằm ở hai
    ô khác nhau thành hai dòng riêng. Ở đây các từ có cùng tọa độ y được
    gom lại một dòng; khoảng trống lớn giữa hai từ được thay bằng 3 dấu
    cách để đánh dấu ranh giới cột.

    words: danh sách (x0, y0, x1, y1, text, ...) như page.get_text("words").
    """
    items = sorted(
        (w for w in words if str(w[4]).strip()),
        key=lambda w: ((w[1] + w[3]) / 2, w[0])
    )

    lines = []
    for w in items:
        y_center = (w[1] + w[3]) / 2
        if lines and abs(lines[-1]["y"] - y_center) <= y_tolerance:
            lines[-1]["words"].append(w)
        else:
            lines.append({"y": y_center, "words": [w]})

    text_lines = []
    for line in lines:
        line_words = sorted(line["words"], key=lambda w: w[0])
        parts = [str(line_words[0][4])]

        for prev, cur in zip(line_words, line_words[1:]):
            height = max(prev[3] - prev[1], 1)
            gap = cur[0] - prev[2]
            # Khoảng trống lớn hơn ~1.2 lần chiều cao chữ = sang cột khác
            parts.append("   " if gap > height * 1.2 else " ")
            parts.append(str(cur[4]))

        text_lines.append("".join(parts))

    return "\n".join(text_lines)


def sgcl_page_tables(page):
    """Bảng có kẻ ô trên trang: grid[row][col] = (bbox, text) hoặc None (ô gộp)."""
    grids = []
    if not hasattr(page, "find_tables"):
        return grids
    try:
        found = page.find_tables()
    except Exception:
        return grids
    for table in found.tables:
        try:
            data = table.extract() or []
            grid = []
            for r, row in enumerate(table.rows):
                cells = []
                for c, bbox in enumerate(row.cells):
                    text = data[r][c] if r < len(data) and c < len(data[r]) else None
                    cells.append((tuple(bbox), str(text or "")) if bbox is not None else None)
                grid.append(cells)
            grids.append(grid)
        except Exception:
            continue
    return grids


def sgcl_get_pdf_pages(file_bytes):
    """Đọc text layer của tất cả các trang PDF: text theo dòng, tọa độ từng từ, bảng."""
    pages = []

    with fitz.open(stream=file_bytes, filetype="pdf") as pdf:
        for index, page in enumerate(pdf, start=1):
            raw_words = page.get_text("words") or []
            text = sgcl_words_to_rows(raw_words) if raw_words else ""

            pages.append({
                "page": index,
                "text": text,
                "words": [tuple(w[:5]) for w in raw_words],
                "tables": sgcl_page_tables(page) if raw_words else [],
            })

    return pages


def sgcl_tesseract_languages():
    """Chọn ngôn ngữ OCR có sẵn; không giả định máy chủ đã cài tiếng Việt."""
    import pytesseract
    try:
        available = set(pytesseract.get_languages(config=""))
    except Exception:
        available = {"eng"}
    selected = [lang for lang in ("vie", "eng") if lang in available]
    return "+".join(selected) if selected else "eng"


def sgcl_remove_table_lines(image):
    """
    Xóa đường kẻ ngang/dọc dài (khung bảng) để OCR không bỏ sót chữ trong ô.
    Xử lý từng trục với kiểu dữ liệu nhỏ để giữ RAM thấp.
    """
    import numpy as np

    gray = np.asarray(image.convert("L"))
    dark = gray < 160
    del gray
    h, w = dark.shape

    def line_mask(mask, length):
        """Đánh dấu điểm thuộc đoạn tối liên tục dài >= length theo trục ngang."""
        n = mask.shape[1]
        if length < 2 or n < length:
            return np.zeros(mask.shape, dtype=bool)
        csum = np.zeros((mask.shape[0], n + 1), dtype=np.int32)
        np.cumsum(mask, axis=1, dtype=np.int32, out=csum[:, 1:])
        # span[i] = có cửa sổ dài `length` toàn điểm tối bắt đầu tại i
        span = np.zeros(mask.shape, dtype=bool)
        span[:, :n - length + 1] = (csum[:, length:] - csum[:, :-length]) == length
        del csum
        # Điểm j thuộc đường nếu có cửa sổ bắt đầu trong [j-length+1, j]
        covered = 1
        while covered < length:
            step = min(covered, length - covered)
            span[:, step:] |= span[:, :-step].copy()
            covered += step
        return span

    horizontal = line_mask(dark, max(80, int(w * 0.06)))
    vertical = line_mask(dark.T, max(60, int(h * 0.02))).T
    del dark
    lines = horizontal | vertical
    del horizontal, vertical
    if not lines.any():
        return image

    # nới rộng 1 px để xóa cả viền mờ
    grown = lines.copy()
    grown[1:, :] |= lines[:-1, :]
    grown[:-1, :] |= lines[1:, :]
    grown[:, 1:] |= lines[:, :-1]
    grown[:, :-1] |= lines[:, 1:]
    del lines
    out = np.array(image.convert("RGB"))
    out[grown] = 255
    return Image.fromarray(out)


def sgcl_ocr_page(file_bytes, page_number, dpi_scale=3):
    """
    OCR một trang. Trả về (text, words) với words theo tọa độ trang PDF,
    để dùng chung bộ trích xuất theo vị trí như text layer.
    """
    import pytesseract
    from pytesseract import Output

    with fitz.open(stream=file_bytes, filetype="pdf") as pdf:
        page = pdf[page_number - 1]
        # ~3300 px cạnh dài ≈ 300 dpi khổ A4: đủ nét cho OCR, không quá tải RAM
        dpi_scale = sgcl_safe_zoom(page, dpi_scale, max_side=3300)
        pix = page.get_pixmap(matrix=fitz.Matrix(dpi_scale, dpi_scale), alpha=False)
        image = Image.open(BytesIO(pix.tobytes("png"))).convert("RGB")

    try:
        image = sgcl_remove_table_lines(image)
    except Exception:
        pass

    data = pytesseract.image_to_data(
        image, lang=sgcl_tesseract_languages(), config="--oem 3 --psm 6",
        output_type=Output.DICT
    )
    words = []
    for i, raw in enumerate(data.get("text", [])):
        word = str(raw or "").strip()
        try:
            confidence = float(data["conf"][i])
        except Exception:
            confidence = -1
        # Bỏ ký tự rác do đường kẻ còn sót
        if not word or confidence < 0 or re.fullmatch(r"[|\[\]_—–\-]+", word):
            continue
        left, top = int(data["left"][i]), int(data["top"][i])
        width, height = int(data["width"][i]), int(data["height"][i])
        words.append((
            left / dpi_scale, top / dpi_scale,
            (left + width) / dpi_scale, (top + height) / dpi_scale, word
        ))

    return sgcl_words_to_rows(words), words


def sgcl_ocr_page_text(file_bytes, page_number, dpi_scale=3):
    """Giữ tương thích: chỉ trả về text OCR."""
    return sgcl_ocr_page(file_bytes, page_number, dpi_scale)[0]


def sgcl_get_all_pdf_text(file_bytes):
    """Đọc text layer nếu đủ chất lượng; OCR riêng từng trang scan/thiếu chữ."""
    pages = sgcl_get_pdf_pages(file_bytes)
    for item in pages:
        text = item.get("text", "")
        visible = re.sub(r"\s+", "", text)
        alpha_num = sum(ch.isalnum() for ch in visible)
        # Không chỉ kiểm tra độ dài: text layer lỗi/thiếu cũng cần OCR.
        needs_ocr = len(visible) < 35 or alpha_num < max(12, len(visible) * 0.35)
        if needs_ocr:
            try:
                item["text"], item["words"] = sgcl_ocr_page(file_bytes, item["page"])
                item["tables"] = []
                item["source_type"] = "OCR"
            except Exception as e:
                raise RuntimeError(
                    f"Không thể OCR trang {item['page']}. Hãy cài Tesseract OCR "
                    f"và gói ngôn ngữ eng/vie. Chi tiết: {e}"
                ) from e
        else:
            item["source_type"] = "Text layer"
    return pages


def sgcl_label_pattern(label):
    """
    Tạo regex khớp nhãn theo từ nguyên vẹn, bỏ qua khác biệt dấu câu/khoảng
    trắng. VD "B/L No." khớp "B/L No", "B / L  NO." nhưng "No." không khớp
    bên trong chữ "Notice".
    """
    tokens = re.findall(r"[A-Za-z0-9]+", str(label))
    if not tokens:
        return ""
    body = r"[\W_]*".join(re.escape(t) for t in tokens)
    return rf"(?<![A-Za-z0-9]){body}(?![A-Za-z0-9])"


def sgcl_value_after_label(rest):
    """
    Lấy giá trị ngay sau nhãn: bỏ dấu ':' '.' '-' đứng đầu, rồi chỉ giữ
    cột đầu tiên (ngăn bởi 3 dấu cách) để không dính sang nhãn của cột khác.
    """
    rest = rest.lstrip(" \t:：.-–—)]")
    segments = [s.strip() for s in re.split(r"\s{3,}", rest) if s.strip()]
    if not segments:
        return ""
    return re.sub(r"\s+", " ", segments[0]).strip(" :：")

# =========================================================
# TRÍCH XUẤT TRƯỜNG ĐƠN THEO TỌA ĐỘ CHỮ (text layer hoặc OCR)
# =========================================================
# Mỗi từ là tuple (x0, y0, x1, y1, text) theo tọa độ trang PDF.
# Bảng có kẻ ô (nếu có) là danh sách lưới: grid[row][col] = (bbox, text) hoặc None.

SGCL_FIELD_ALIASES = {
    "IMO Number": ["IMO No", "IMO No.", "IMO number", "IMO"],
    "IMO No.": ["IMO Number", "IMO No", "IMO"],
    "Ship name": ["Name of ship", "Vessel name", "Name of vessel"],
    "Call sign": ["Callsign", "Radio call sign"],
    "Arrival draft (F/A)": ["Arrival draft", "Draft F/A", "Arrival draught"],
    "Air draft": ["Air draft", "Air draught"],
    "Last port of call": ["Last port of call", "Last port"],
    "Port of arrival": ["Port of arrival", "Arrival port"],
    "Departure from last port": ["Departure from last port"],
    "Time of arrival": ["Time of arrival", "Arrival time"],
    "All fast": ["All fast", "Allfast"],
    "P.O.B": ["P.O.B", "Pilot on board"],
    "F.O": ["Fuel oil", "F.O", "FO", "HFO"],
    "M.D.O": ["Diesel oil", "M.D.O", "MDO", "MGO"],
    "FreshWater": ["Fresh water", "FreshWater", "F.W"],
    "Berthing Position": ["Berthing position", "Berth"],
    "Port of discharge": ["Port of discharge"],
    "B/L No.": ["B/L No", "Bill of lading No", "B/L Number"],
    "Gross weight": ["Gross weight"],
    "Port of Loading": ["Port of loading"],
    "Port of Discharge": ["Port of discharge"],
    "Arriving from": ["Arriving from"],
    "Sailing to": ["Sailing to"],
    "Issued at": ["Issued at"],
    "Number of persons on board": ["Number of persons on board", "Persons on board"],
    "Period of stay": ["Period of stay"],
    "This certificate is valid until": ["This certificate is valid until", "valid until"],
}

# Trường có nội dung nhiều dòng (khối địa chỉ, mô tả hàng)
SGCL_BLOCK_FIELDS = {
    "shipper", "consignee", "notifyparty",
    "shippersdescriptionofgoods", "descriptionofgoods",
}


def sgcl_label_tokens(label):
    return re.findall(r"[A-Za-z0-9]+", str(label or ""))


def sgcl_field_labels(field_name):
    """Các cách viết nhãn của một trường, nhãn dài thử trước."""
    labels = [field_name] + SGCL_FIELD_ALIASES.get(field_name, [])
    # "Arrival draft (F/A)" -> thêm "Arrival draft"; "Re-inspection required?" -> bỏ "?"
    base = re.sub(r"\s*\([^)]*\)\s*$", "", str(field_name)).strip()
    if base:
        labels.append(base)
    unique = {}
    for label in labels:
        tokens = sgcl_label_tokens(label)
        if tokens:
            unique.setdefault(" ".join(t.lower() for t in tokens), label)
    return sorted(unique.values(), key=lambda x: len(sgcl_label_tokens(x)), reverse=True)


_SGCL_OCR_CONFUSION = {"o": "[o0]", "i": "[il1|]", "l": "[il1|]", "s": "[s5]"}


def _sgcl_fuzzy_token(token, loose=False):
    """
    Cho phép lỗi OCR hay gặp: O<->0, I/l<->1, S<->5.
    loose=True: chấp nhận thêm 1 ký tự bị đọc sai (VD 'drafl', 'certiticate').
    """
    if token.isdigit():
        return re.escape(token)

    def char_class(ch):
        return _SGCL_OCR_CONFUSION.get(ch.lower(), re.escape(ch))

    if len(token) == 1:
        # Chữ đơn lẻ (P.O.B, M.D.O) - OCR hay nhân đôi, VD 'P.0O.B'
        return char_class(token) + ("{1,2}" if loose else "")
    exact = "".join(char_class(ch) for ch in token)
    if not loose or len(token) < 4:
        return exact
    variants = [exact] + [
        "".join(char_class(ch) if j != i else "[A-Za-z0-9]" for j, ch in enumerate(token))
        for i in range(len(token))
    ]
    return "(?:" + "|".join(variants) + ")"


def _sgcl_label_body(label, loose=False):
    tokens = sgcl_label_tokens(label)
    if not tokens:
        return ""
    # Nhãn 1 từ ngắn (VD 'Port', 'Unit') không nới lỏng để tránh bắt nhầm
    loose = loose and (len(tokens) >= 2 or len(tokens[0]) >= 6)
    return r"[\W_]*".join(_sgcl_fuzzy_token(t, loose) for t in tokens)


def _sgcl_label_regex(label, loose=False):
    body = _sgcl_label_body(label, loose)
    return re.compile(rf"(?<![A-Za-z0-9]){body}(?![A-Za-z0-9])", re.I) if body else None


def _sgcl_ocr_key(text):
    """Chuỗi so sánh gần đúng: bỏ dấu câu, quy 0->o, 1->l, 5->s."""
    key = sgcl_normalize(text)
    return key.translate(str.maketrans({"0": "o", "1": "l", "5": "s", "i": "l"}))


class SgclLabelSet:
    """Tập nhãn đã biết của loại chứng từ, dùng để loại giá trị là nhãn."""

    def __init__(self, labels):
        self.items = []
        seen = set()
        for label in labels:
            tokens = sgcl_label_tokens(label)
            key = " ".join(t.lower() for t in tokens)
            if not tokens or key in seen:
                continue
            seen.add(key)
            body = _sgcl_label_body(label)
            self.items.append({
                "key": _sgcl_ocr_key(label),
                "tokens": len(tokens),
                "full": re.compile(rf"^[\W_]*{body}[\W_]*$", re.I),
                "search": re.compile(rf"(?<![A-Za-z0-9]){body}(?![A-Za-z0-9])", re.I),
            })

    def is_label(self, text):
        text = str(text or "").strip()
        if not text:
            return False
        if self.starts_with_label(text):
            return True
        first = text.splitlines()[0].strip()
        for item in self.items:
            if item["full"].match(text) or item["full"].match(first):
                return True
        # Nhãn bị OCR đọc sai nhẹ, VD 'All tast', 'P.0O.B', 'M.D.OD'
        letters = sum(ch.isalpha() for ch in first)
        if letters and letters >= len(re.sub(r"\W", "", first)) * 0.6 and len(first) <= 60:
            from difflib import SequenceMatcher
            key = _sgcl_ocr_key(first)
            for item in self.items:
                if len(item["key"]) >= 3 and SequenceMatcher(None, key, item["key"]).ratio() >= 0.8:
                    return True
        # "Something:" ngắn, không có số -> là nhãn chưa cấu hình
        if re.search(r"[:：]\s*$", first) and not re.search(r"\d", first) and len(first) <= 60:
            return True
        return False

    def starts_with_label(self, text):
        """Ứng viên mở đầu bằng nhãn khác, VD 'B/L No.: HTPO...' khi đang tìm Shipper."""
        text = str(text or "").strip()
        for item in self.items:
            m = item["search"].match(text)
            if not m:
                continue
            rest = text[m.end():]
            if item["tokens"] >= 2 or re.match(r"\s*[:：]", rest) or not rest.strip():
                return True
        return False

    def cut_at_next_label(self, text):
        """Cắt giá trị khi gặp nhãn khác phía sau (VD 'IMO 123 Call sign ABC')."""
        cut = len(text)
        for item in self.items:
            for m in item["search"].finditer(text):
                if m.start() == 0:
                    continue
                followed_by_colon = bool(re.match(r"\s*[:：]", text[m.end():]))
                if item["tokens"] >= 2 or followed_by_colon:
                    cut = min(cut, m.start())
                break
        return text[:cut].strip()


def sgcl_clean_value(text, label_set=None):
    """Làm sạch giá trị: bỏ dấu ':' đầu, phần dịch nhãn '/ Tên tàu:', ngoặc chú thích."""
    lines = []
    for raw in str(text or "").splitlines():
        line = re.sub(r"[ \t]+", " ", raw).strip()
        if line:
            lines.append(line)
    text = "\n".join(lines)

    for _ in range(3):
        before = text
        text = text.lstrip(" \t\n:：.-–—)]|")
        # Ngoặc chú thích ngay sau nhãn, VD "(F/A):" hoặc "(Người gửi hàng)"
        text = re.sub(r"^\([^)\d]{0,40}\)\s*", "", text)
        # Phần nhãn song ngữ trước dấu ':' VD "/ Tên tàu: HTP OCEAN"
        m = re.match(r"^([^:：\n\d]{0,45})[:：]\s*(.+)$", text, re.S)
        if m and (m.group(1).lstrip().startswith(("/", "(")) or
                  any(ord(ch) > 127 for ch in m.group(1))):
            text = m.group(2)
        if text == before:
            break

    # OCR hay tách số trong ngày tháng: "20/1 1/2026" -> "20/11/2026"
    text = re.sub(r"\b(\d)\s(\d)(?=/\d{1,2}/\d{2,4}\b)", r"\1\2", text)
    text = re.sub(r"\b(\d{1,2})/(\d)\s(\d)/(\d{2,4})\b", r"\1/\2\3/\4", text)

    if label_set is not None and text:
        first, *others = text.split("\n")
        first = label_set.cut_at_next_label(first)
        text = "\n".join([first] + others).strip()

    return text.strip(" \t\n:：|")


# ---------------------------------------------------------
# GOM TỪ THÀNH DÒNG VÀ ĐOẠN (CỘT)
# ---------------------------------------------------------

def sgcl_build_lines(words):
    items = sorted(
        (w for w in words if str(w[4]).strip()),
        key=lambda w: ((w[1] + w[3]) / 2, w[0])
    )
    lines = []
    for w in items:
        yc = (w[1] + w[3]) / 2
        h = max(w[3] - w[1], 1.0)
        if lines and abs(lines[-1]["yc"] - yc) <= max(2.5, h * 0.45):
            lines[-1]["words"].append(w)
        else:
            lines.append({"yc": yc, "words": [w]})

    for line in lines:
        ws = sorted(line["words"], key=lambda w: w[0])
        heights = sorted(max(w[3] - w[1], 1.0) for w in ws)
        line["words"] = ws
        line["h"] = heights[len(heights) // 2]
        line["y0"] = min(w[1] for w in ws)
        line["y1"] = max(w[3] for w in ws)
        text, spans, pos = [], [], 0
        for w in ws:
            t = str(w[4])
            spans.append((pos, pos + len(t)))
            text.append(t)
            pos += len(t) + 1
        line["text"] = " ".join(text)
        line["spans"] = spans
        # Đoạn = nhóm từ liền nhau, tách bởi khoảng trống lớn
        segs, start = [], 0
        for k in range(1, len(ws)):
            gap = ws[k][0] - ws[k - 1][2]
            if gap > max(line["h"] * 1.2, 6):
                segs.append((start, k))
                start = k
        segs.append((start, len(ws)))
        line["segs"] = segs
    return lines


def _seg_of(line, k):
    for s, e in line["segs"]:
        if s <= k < e:
            return s, e
    return 0, len(line["words"])


def _join_words(ws):
    return " ".join(str(w[4]) for w in ws)


def _find_label_hits(lines, label, loose=False):
    """Tìm vị trí nhãn; hỗ trợ nhãn dài bị xuống dòng trong cùng một cột."""
    regex = _sgcl_label_regex(label, loose)
    if regex is None:
        return []
    hits = []
    for li, line in enumerate(lines):
        for m in regex.finditer(line["text"]):
            idx = [k for k, (s, e) in enumerate(line["spans"]) if s < m.end() and e > m.start()]
            if not idx:
                continue
            last = idx[-1]
            tail = line["text"][m.end():line["spans"][last][1]]
            hits.append({
                "words": [(li, k) for k in idx],
                "right": [(li, last + 1, tail)],
                "first_line": li,
            })
    if hits or len(sgcl_label_tokens(label)) < 3:
        return hits

    # Nhãn dài xuống 2-3 dòng trong cùng cột
    for li, line in enumerate(lines):
        for s, e in line["segs"]:
            seg_words = line["words"][s:e]
            x0 = seg_words[0][0]
            next_x = line["words"][e][0] if e < len(line["words"]) else float("inf")
            h = line["h"]
            combined = [(li, k) for k in range(s, e)]
            first_len = len(_join_words(seg_words))
            prev_y1 = line["y1"]
            for lj in (li + 1, li + 2):
                if lj >= len(lines) or lines[lj]["y0"] - prev_y1 > h * 1.2:
                    break
                cont = [k for k, w in enumerate(lines[lj]["words"])
                        if w[0] >= x0 - h and w[2] <= next_x - 1]
                if not cont:
                    break
                cont_seg = _seg_of(lines[lj], cont[0])
                cont = [k for k in cont if k < cont_seg[1]]
                combined += [(lj, k) for k in cont]
                prev_y1 = lines[lj]["y1"]
                text = _join_words(lines[a]["words"][b] for a, b in combined)
                m = regex.search(text)
                if not m or m.start() >= first_len or m.end() <= first_len:
                    continue
                pos, used = 0, []
                for a, b in combined:
                    t = str(lines[a]["words"][b][4])
                    if pos < m.end() and pos + len(t) > m.start():
                        used.append((a, b, pos + len(t)))
                    pos += len(t) + 1
                la, lb, end_pos = used[-1]
                tail = text[m.end():end_pos]
                hits.append({
                    "words": [(a, b) for a, b, _ in used],
                    "right": [(la, lb + 1, tail), (li, e, "")],
                    "first_line": li,
                })
                break
    return hits


def _right_value(lines, li, start, tail, label_set):
    """
    Giá trị nằm bên phải nhãn trên cùng dòng.
    Trả về (giá_trị, tách_rõ): tách_rõ = có ':' '?' hoặc khoảng trống cột giữa
    nhãn và giá trị. Không tách rõ thường là nhãn nằm trong câu văn.
    """
    line = lines[li]
    if tail and re.search(r"[A-Za-z0-9]", tail):
        # Nhãn dính liền giá trị, VD "IMO:9123456"
        s, e = _seg_of(line, max(start - 1, 0))
        text = tail + " " + _join_words(line["words"][start:e])
        value = sgcl_clean_value(text, label_set)
        separated = bool(re.match(r"^[^A-Za-z0-9]*[:：?=]", tail))
        return (value, separated) if value and not label_set.is_label(value) else ("", False)

    separated = bool(re.search(r"[:：?=]", tail or ""))
    # Bỏ các từ chỉ là dấu câu ':' '-' '>' đứng đầu
    k = start
    while k < len(line["words"]) and not re.search(r"[A-Za-z0-9\u00C0-\u1EF9]", str(line["words"][k][4])):
        separated = True
        k += 1
    if k >= len(line["words"]):
        return "", False
    s, e = _seg_of(line, k)
    if start == 0 or s == k and k > 0 and _seg_of(line, start - 1) != (s, e):
        separated = True
    value = sgcl_clean_value(_join_words(line["words"][k:e]), label_set)
    if not value or label_set.is_label(value):
        return "", False
    return value, separated


def _below_value(lines, hit, label_set, multiline):
    used = hit["words"]
    first = lines[hit["first_line"]]
    lx0 = min(lines[a]["words"][b][0] for a, b in used)
    lx1 = max(lines[a]["words"][b][2] for a, b in used)
    bottom = max(lines[a]["words"][b][3] for a, b in used)
    h = first["h"]

    # Biên cột: nửa khoảng cách tới đoạn bên trái/phải trên dòng nhãn
    first_idx = [b for a, b in used if a == hit["first_line"]]
    s, e = _seg_of(first, first_idx[0])
    left = first["words"][s - 1][2] if s > 0 else None
    right = first["words"][e][0] if e < len(first["words"]) else None
    col_l = (left + lx0) / 2 if left is not None else lx0 - 3 * h
    col_r = (lx1 + right) / 2 if right is not None else float("inf")

    def in_col(w):
        width = max(w[2] - w[0], 0.1)
        overlap = min(w[2], col_r) - max(w[0], col_l)
        return overlap >= width * 0.5

    collected, skipped, last_y1 = [], 0, bottom
    last_line = max(a for a, _ in used)
    for lj in range(last_line + 1, len(lines)):
        line = lines[lj]
        if line["y0"] - last_y1 > h * (1.6 if collected else 4.5):
            break
        ws = [w for w in line["words"] if in_col(w)]
        if not ws:
            if collected:
                break
            continue
        text = sgcl_clean_value(_join_words(ws), label_set)
        if not text or label_set.is_label(text):
            if collected:
                break
            skipped += 1
            if skipped > 2:
                break
            last_y1 = line["y1"]
            continue
        collected.append(text)
        last_y1 = line["y1"]
        if not multiline or len(collected) >= 6:
            break
    return "\n".join(collected)


def _cell_below(grid, r, bbox):
    xc = (bbox[0] + bbox[2]) / 2
    for r2 in range(r + 1, len(grid)):
        for cell in grid[r2]:
            if cell and cell[0][0] - 1 <= xc <= cell[0][2] + 1:
                return cell
    return None


def _value_from_tables(tables, labels, label_set):
    for grid in tables or []:
        for r, row in enumerate(grid):
            for c, cell in enumerate(row):
                if not cell or not str(cell[1]).strip():
                    continue
                bbox, text = cell
                for label in labels:
                    regex = _sgcl_label_regex(label)
                    m = regex.search(text) if regex else None
                    if not m:
                        continue
                    if sum(ch.isalnum() for ch in text[:m.start()]) > 30:
                        continue
                    rest = sgcl_clean_value(text[m.end():], label_set)
                    if rest and not label_set.is_label(rest):
                        return rest
                    for c2 in range(c + 1, len(row)):
                        if row[c2] and str(row[c2][1]).strip():
                            cand = sgcl_clean_value(row[c2][1], label_set)
                            if cand and not label_set.is_label(cand):
                                return cand
                            break
                    below = _cell_below(grid, r, bbox)
                    if below:
                        cand = sgcl_clean_value(below[1], label_set)
                        if cand and not label_set.is_label(cand):
                            return cand
                    break
    return ""


def sgcl_extract_single(pages, field_name, other_fields=None):
    """
    Lấy GIÁ TRỊ của trường (không phải tên trường). Các vị trí được thử:
    - Ô bảng chứa nhãn -> phần còn lại của ô / ô bên phải / ô bên dưới.
    - Cùng dòng, sau nhãn (VD 'Air draft: 12.5 m').
    - Ngay bên dưới nhãn, cùng cột (hàng tiêu đề - hàng giá trị).
    Mọi ứng viên trùng với một nhãn đã biết đều bị loại. Khi có nhiều ứng
    viên, ưu tiên dạng biểu mẫu rõ ràng hơn nhãn nằm lẫn trong câu văn.
    """
    labels = sgcl_field_labels(field_name)
    all_labels = list(labels)
    for name in other_fields or []:
        all_labels += sgcl_field_labels(name)
    label_set = SgclLabelSet(all_labels)
    multiline = sgcl_normalize(field_name) in SGCL_BLOCK_FIELDS

    best = None  # (điểm, thứ tự, giá trị, trang)
    order = 0

    def consider(score, value, page_no):
        nonlocal best, order
        order += 1
        if value and (best is None or score > best[0]):
            best = (score, order, value, page_no)

    for page in pages:
        value = _value_from_tables(page.get("tables"), labels, label_set)
        if value:
            consider(4.5, value, page["page"])

        words = page.get("words")
        if words is None:
            words = sgcl_words_from_text(page.get("text", ""))
        lines = page.get("_lines")
        if lines is None:
            lines = sgcl_build_lines(words)
            page["_lines"] = lines

        for label in labels:
            seen = set()
            for loose in (False, True):
                bonus = 0.0 if loose else 0.5
                for hit in _find_label_hits(lines, label, loose):
                    key = tuple(hit["words"])
                    if key in seen:
                        continue
                    seen.add(key)

                    right, separated = "", False
                    for li, start, tail in hit["right"]:
                        right, separated = _right_value(lines, li, start, tail, label_set)
                        if right:
                            break

                    if right and separated and not multiline:
                        consider(3 + bonus, right, page["page"])
                        continue

                    below = _below_value(lines, hit, label_set, multiline)
                    if multiline and (right or below):
                        value = "\n".join(x for x in (right, below) if x)
                        consider((3 if separated or below else 1) + bonus, value, page["page"])
                    elif below:
                        consider(2 + bonus, below, page["page"])
                    elif right:
                        consider(1 + bonus, right, page["page"])

    if best:
        return {"value": best[2], "page": best[3], "found": True}
    return {"value": "", "page": None, "found": False}


def sgcl_words_from_text(text):
    """Dự phòng khi trang chỉ có text: tạo tọa độ giả theo dòng/ký tự."""
    words = []
    for li, line in enumerate(str(text or "").splitlines()):
        for m in re.finditer(r"\S+", line):
            words.append((m.start() * 5.0, li * 12.0, m.end() * 5.0, li * 12.0 + 10, m.group()))
    return words


def sgcl_match_columns(table_header, headers, min_score=0.55):
    """
    Ghép từng trường bảng đã cấu hình với cột tương ứng của bảng trong PDF.
    Trả về danh sách chỉ số cột (None nếu không ghép được).
    """
    from difflib import SequenceMatcher

    normalized_cells = [sgcl_normalize(c) for c in table_header]
    result = []
    used = set()

    for header in headers:
        key = sgcl_normalize(header)
        best_index, best_score = None, 0.0

        for index, cell in enumerate(normalized_cells):
            if index in used or not cell or not key:
                continue
            if key in cell or cell in key:
                score = 0.9 + min(len(cell), len(key)) / max(len(cell), len(key)) / 10
            else:
                score = SequenceMatcher(None, key, cell).ratio()
            if score > best_score:
                best_index, best_score = index, score

        if best_index is not None and best_score >= min_score:
            used.add(best_index)
            result.append(best_index)
        else:
            result.append(None)

    return result


def sgcl_is_header_row(row, headers):
    """Dòng toàn tên cột (tiêu đề lặp lại ở trang sau) không phải dữ liệu."""
    cells = [str(c or "").strip() for c in row if str(c or "").strip()]
    if not cells:
        return False
    label_set = SgclLabelSet(headers)
    hits = sum(1 for c in cells if label_set.is_label(c))
    return hits * 2 > len(cells)


def sgcl_extract_table_from_pdf(file_bytes, headers):
    """
    Dùng bộ nhận diện bảng của PyMuPDF (page.find_tables) cho bảng có kẻ ô
    như Crew List, Ship's Stores List. Chọn bảng có tiêu đề khớp nhiều
    trường cấu hình nhất, gộp các trang nếu bảng kéo dài nhiều trang.
    """
    collected = []
    first_page = None

    with fitz.open(stream=file_bytes, filetype="pdf") as pdf:
        for page_index, page in enumerate(pdf, start=1):
            if not hasattr(page, "find_tables"):
                return None
            try:
                found = page.find_tables()
            except Exception:
                continue

            for table in found.tables:
                data = table.extract() or []
                if not data:
                    continue

                # Tiêu đề có thể nằm ở 1 trong 3 dòng đầu
                best = None
                for header_row in range(min(3, len(data))):
                    cells = [str(c or "").replace("\n", " ") for c in data[header_row]]
                    col_index = sgcl_match_columns(cells, headers)
                    matched = sum(i is not None for i in col_index)
                    if best is None or matched > best[0]:
                        best = (matched, header_row, col_index)

                matched, header_row, col_index = best
                minimum_matches = max(2, (len(headers) * 3 + 4) // 5)
                if matched < minimum_matches:
                    continue
                if headers and "familyname" in sgcl_normalize(headers[0]) and matched < len(headers):
                    continue
                # Crew List phải nhận diện được cột họ tên; tránh lấy nhầm
                # một dòng tiêu đề phụ làm header rồi đẩy dữ liệu lệch cột.
                if headers and "familyname" in sgcl_normalize(headers[0]):
                    if not col_index or col_index[0] is None:
                        continue

                for raw in data[header_row + 1:]:
                    row = [
                        re.sub(r"\s+", " ", str(raw[i] or "")).strip()
                        if i is not None and i < len(raw) else ""
                        for i in col_index
                    ]
                    if any(row) and not sgcl_is_header_row(row, headers):
                        collected.append(row)

                if first_page is None:
                    first_page = page_index

    if not collected:
        return None

    return {
        "headers": headers,
        "rows": collected,
        "page": first_page,
        "found": True
    }


def sgcl_extract_table(pages, headers, file_bytes=None):
    """
    Tách bảng từ PDF.
    1. Thử nhận diện bảng có kẻ ô bằng PyMuPDF.
    2. Nếu không được, tìm dòng tiêu đề trong text rồi lấy các dòng phía sau.
    Không tự suy đoán hoặc tạo dữ liệu khi không tìm thấy.
    """
    empty = {"headers": headers or [], "rows": [], "page": None, "found": False}

    if not headers:
        return empty

    # Bảng Crew List trên bản scan mẫu có nhiều cột sát nhau và OCR dễ
    # đảo cột. Không tự tạo bảng sai cấu trúc: giữ kết quả trống để người
    # dùng kiểm tra/nhập lại trong bảng chỉnh sửa ở Bước 3.
    if ("familyname" in sgcl_normalize(headers[0])
            and any(page.get("source_type") == "OCR" for page in pages)):
        return empty

    if file_bytes is not None:
        try:
            result = sgcl_extract_table_from_pdf(file_bytes, headers)
            if result:
                return result
        except Exception:
            pass

    rows = []
    source_page = None
    min_matched = max(2, (len(headers) * 3 + 4) // 5)

    for page in pages:
        lines = page["text"].splitlines()
        col_index = None
        empty_streak = 0

        for line in lines:
            parts = [
                re.sub(r"\s+", " ", x).strip()
                for x in re.split(r"\t+|\s{2,}", line.strip())
                if x.strip()
            ]

            # Chưa gặp dòng tiêu đề: chỉ tìm tiêu đề, bỏ qua mọi dòng khác
            if col_index is None:
                if len(parts) >= 2:
                    candidate = sgcl_match_columns(parts, headers)
                    enough_matches = sum(i is not None for i in candidate) >= min_matched
                    first_column_ok = not (
                        headers and "familyname" in sgcl_normalize(headers[0])
                    ) or bool(candidate and candidate[0] is not None)
                    all_crew_columns = not (
                        headers and "familyname" in sgcl_normalize(headers[0])
                    ) or sum(i is not None for i in candidate) == len(headers)
                    if enough_matches and first_column_ok and all_crew_columns:
                        col_index = candidate
                continue

            if len(parts) < 2:
                empty_streak += 1
                # Hết bảng khi gặp nhiều dòng liên tiếp không có dạng cột
                if empty_streak >= 3:
                    col_index = None
                    empty_streak = 0
                continue

            empty_streak = 0

            if len(parts) > len(headers) + 2:
                continue

            # Số cột của dòng dữ liệu thường khớp vị trí với dòng tiêu đề
            row = [
                parts[i] if i is not None and i < len(parts) else ""
                for i in col_index
            ]

            if any(row) and not sgcl_is_header_row(row, headers):
                rows.append(row)
                if source_page is None:
                    source_page = page["page"]

    return {
        "headers": headers,
        "rows": rows,
        "page": source_page,
        "found": bool(rows)
    }

def sgcl_ensure_ship_data(ship):
    ship.setdefault("field_extractions", [])
    ship.setdefault("mappings", [])
    ship.setdefault("manual_entries", [])
    ship.setdefault("review_confirmed", False)
    ship.setdefault("excel_output_bytes", None)
    ship.setdefault("step_status", {})


def sgcl_invalidate_review(ship):
    ship["review_confirmed"] = False
    ship["excel_output_bytes"] = None
    ship["step_status"]["4"] = "pending"
    ship["step_status"]["5"] = "pending"

    # Dữ liệu đã thay đổi thì hồ sơ không còn ở trạng thái xác nhận/hoàn thành
    if ship.get("status") in ("Đã xác nhận", "Hoàn thành"):
        ship["status"] = "Đã trích xuất"


# =========================================================
# GHI NHỚ VỊ TRÍ ÁNH XẠ EXCEL GIỮA CÁC HỒ SƠ
# =========================================================

MAPPING_MEMORY_PATH = Path(__file__).resolve().parent / "sgcl_mapping_memory.json"


def sgcl_mapping_key(doc_type, field):
    return f"{sgcl_normalize(doc_type)}||{sgcl_normalize(field)}"


def sgcl_default_mapping():
    """Ánh xạ mặc định theo TemplateNSW.xlsx."""
    result = {}
    for doc_type, field, sheet, cell, extra, column_map in DEFAULT_NSW_MAPPING_LIST:
        result[sgcl_mapping_key(doc_type, field)] = {
            "sheet": sheet,
            "cell": cell,
            "extra_targets": [{"sheet": s, "cell": c} for s, c in extra],
            "column_map": dict(column_map or {}),
        }
    return result


def sgcl_load_mapping_memory():
    """Vị trí đã dùng ở các hồ sơ trước, bổ sung bằng ánh xạ mặc định."""
    memory = sgcl_default_mapping()
    if MAPPING_MEMORY_PATH.exists():
        try:
            memory.update(
                json.loads(MAPPING_MEMORY_PATH.read_text(encoding="utf-8"))
            )
        except Exception:
            pass
    return memory


def sgcl_format_targets(targets):
    """[{'sheet','cell'}] -> "Sheet!B10; Sheet 2!C5" để hiển thị/sửa."""
    return "; ".join(f"{t['sheet']}!{t['cell']}" for t in targets or [])


def sgcl_parse_targets(text):
    """
    Đọc chuỗi "Sheet!B10; Sheet 2!C5" thành danh sách vị trí.
    Trả về (danh_sách, lỗi).
    """
    targets, errors = [], []
    for part in re.split(r"[;\n]+", text or ""):
        part = part.strip()
        if not part:
            continue
        if "!" not in part:
            errors.append(f"'{part}' thiếu dấu ! (VD: Notice ship to port!B10)")
            continue
        sheet, cell = part.rsplit("!", 1)
        sheet = sheet.strip().strip("'\"")
        cell = cell.strip().upper()
        if sheet not in NSW_SHEETS:
            errors.append(f"Sheet không hợp lệ: '{sheet}'")
            continue
        try:
            openpyxl.utils.cell.coordinate_from_string(cell)
        except Exception:
            errors.append(f"Ô không hợp lệ: '{cell}'")
            continue
        targets.append({"sheet": sheet, "cell": cell})
    return targets, errors


def sgcl_save_mapping_memory(ship):
    """
    Lưu sheet/ô/cột đã chọn cho từng trường để hồ sơ tàu sau
    tự điền sẵn, không phải ánh xạ lại từ đầu.
    """
    memory = sgcl_load_mapping_memory()

    for m in ship.get("mappings", []):
        if not m.get("cell") or not m.get("doc_type"):
            continue
        memory[sgcl_mapping_key(m["doc_type"], m["field"])] = {
            "sheet": m.get("sheet"),
            "cell": m.get("cell"),
            "extra_targets": m.get("extra_targets", []),
            "column_map": m.get("column_map", {})
        }

    # Chỉ lưu phần khác với mặc định để file ghi nhớ gọn
    defaults = sgcl_default_mapping()
    memory = {k: v for k, v in memory.items() if defaults.get(k) != v}

    MAPPING_MEMORY_PATH.write_text(
        json.dumps(memory, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )


def sgcl_normalize_rows(rows, headers):
    """Đưa mọi dòng về đúng số cột, giá trị dạng chuỗi."""
    n = len(headers)
    return [
        ["" if v is None else str(v) for v in (list(row) + [""] * n)[:n]]
        for row in rows or []
    ]


def sgcl_table_editor(key, headers, rows):
    """
    Bảng chỉnh sửa ổn định.

    st.data_editor lưu thao tác sửa/thêm/xóa dòng theo key và áp lại lên
    dữ liệu đầu vào ở mỗi lần chạy. Nếu đưa dữ liệu ĐÃ sửa vào lại, dòng
    thêm mới bị nhân đôi. Vì vậy dữ liệu gốc được giữ cố định trong
    session cho tới khi bảng bị gỡ khỏi màn hình.

    Trả về (rows_mới, có_thay_đổi).
    """
    current = sgcl_normalize_rows(rows, headers)
    base_key = f"{key}__base"

    if key not in st.session_state or base_key not in st.session_state:
        st.session_state[base_key] = current

    edited_df = st.data_editor(
        pd.DataFrame(st.session_state[base_key], columns=headers),
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        key=key
    )

    new_rows = edited_df.fillna("").astype(str).values.tolist()
    # Bỏ các dòng hoàn toàn trống do bấm "thêm dòng" nhưng chưa nhập
    new_rows = [r for r in new_rows if any(str(v).strip() for v in r)]

    return new_rows, new_rows != [
        r for r in current if any(v.strip() for v in r)
    ]


def sgcl_check_template_rule(ws, row, column, value, search_header=False):
    """
    Kiểm tra giá trị theo ghi chú định dạng có sẵn trong ô tiêu đề/ô nhập
    của TemplateNSW (VD 'Định dạng: dd/MM/yyyy HH:mm', 'Số nguyên',
    'Độ dài tối đa: 50'). Trả về chuỗi cảnh báo hoặc None.
    Không tự sửa dữ liệu.
    """
    comment = ws.cell(row=row, column=column).comment
    if comment is None and search_header:
        # Ô dòng dữ liệu của bảng: quy tắc nằm ở ô tiêu đề phía trên
        for r in range(row - 1, max(0, row - 40), -1):
            c = ws.cell(row=r, column=column).comment
            if c is not None:
                comment = c
                break
            if isinstance(ws.cell(row=r, column=column).value, str) and "\n" in str(ws.cell(row=r, column=column).value):
                break

    if comment is None or value is None:
        return None

    note = re.sub(r"\s+", " ", comment.text or "")
    text = str(value).strip()
    if not text:
        return None

    lower = note.lower()
    if "dd/mm/yyyy hh:mm" in lower:
        if not re.fullmatch(r"\d{2}/\d{2}/\d{4} \d{2}:\d{2}", text):
            return "cần dạng dd/MM/yyyy HH:mm"
    elif "dd/mm/yyyy" in lower:
        if not re.fullmatch(r"\d{2}/\d{2}/\d{4}", text):
            return "cần dạng dd/MM/yyyy"

    if "số nguyên" in lower and not re.fullmatch(r"-?\d+", text):
        return "cần là số nguyên"

    if "số thực" in lower and not re.fullmatch(r"-?\d+([.,]\d{1,2})?", text):
        return "cần là số (tối đa 2 chữ số thập phân)"

    max_len = re.search(r"Độ dài tối đa\s*:\s*(\d+)", note, re.I)
    if max_len and len(text) > int(max_len.group(1)):
        return f"dài {len(text)} ký tự, tối đa {max_len.group(1)}"

    return None


def sgcl_restore_x14_validations(wb, template_path):
    """
    openpyxl bỏ qua danh sách thả xuống kiểu mở rộng (x14) khi lưu.
    Đọc lại chúng từ file mẫu và thêm vào dưới dạng chuẩn để file xuất
    giữ đủ dropdown như template.
    """
    import zipfile
    import html
    from openpyxl.worksheet.datavalidation import DataValidation

    with zipfile.ZipFile(template_path) as z:
        workbook_xml = z.read("xl/workbook.xml").decode("utf-8", "ignore")
        rels_xml = z.read("xl/_rels/workbook.xml.rels").decode("utf-8", "ignore")

        targets = {
            m.group(1): m.group(2)
            for m in re.finditer(
                r'<Relationship[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"',
                rels_xml
            )
        }

        for m in re.finditer(
            r'<sheet [^>]*name="([^"]+)"[^>]*r:id="([^"]+)"',
            workbook_xml
        ):
            sheet_name = html.unescape(m.group(1))
            target = targets.get(m.group(2), "").lstrip("/")
            path = target if target.startswith("xl/") else f"xl/{target}"

            if sheet_name not in wb.sheetnames or path not in z.namelist():
                continue

            sheet_xml = z.read(path).decode("utf-8", "ignore")

            for block in re.findall(
                r"<x14:dataValidation .*?</x14:dataValidation>",
                sheet_xml,
                re.S
            ):
                formula = re.search(r"<xm:f>(.*?)</xm:f>", block, re.S)
                sqref = re.search(r"<xm:sqref>(.*?)</xm:sqref>", block, re.S)
                if not formula or not sqref or 'type="list"' not in block:
                    continue

                dv = DataValidation(
                    type="list",
                    formula1=html.unescape(formula.group(1)),
                    allow_blank=True
                )
                dv.add(sqref.group(1))
                wb[sheet_name].add_data_validation(dv)


def sgcl_validate_template_file(path_or_bytes):
    """Trả về danh sách sheet NSW còn thiếu trong file mẫu."""
    source = BytesIO(path_or_bytes) if isinstance(path_or_bytes, bytes) else path_or_bytes
    wb = openpyxl.load_workbook(source, read_only=True)
    names = wb.sheetnames
    wb.close()
    return [n for n in NSW_SHEETS if n not in names]


def sgcl_write_cell(ws, row, column, value):
    """
    Ghi giá trị vào ô Excel. Nếu ô thuộc vùng gộp (merged cell) thì ghi
    vào ô góc trên-trái của vùng, tránh lỗi 'MergedCell ... read-only'.
    """
    for merged in ws.merged_cells.ranges:
        if (
            merged.min_row <= row <= merged.max_row
            and merged.min_col <= column <= merged.max_col
        ):
            row, column = merged.min_row, merged.min_col
            break

    ws.cell(row=row, column=column, value=value)


# Tải cấu hình mặc định và lưu dùng chung
if "sgcl_field_config" not in st.session_state:
    st.session_state.sgcl_field_config = sgcl_load_field_config()

SGCL_FIELD_CONFIG = st.session_state.sgcl_field_config



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
    st.session_state.ships = load_ships_from_db()

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

show_flash()


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

                elif not imo.strip().isdigit() or len(imo.strip()) != 7:
                    st.error("Số IMO phải gồm đúng 7 chữ số.")

                elif any(
                    s["imo"] == imo.strip()
                    for s in st.session_state.ships
                ):
                    st.error("Số IMO đã tồn tại.")

                else:
                    new_ship = {
                        "id": str(uuid.uuid4()),
                        "name": name.strip(),
                        "imo": imo.strip(),
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
                    save_ship_to_db(new_ship)
                    st.session_state.show_form = False

                    flash(f"Đã tạo hồ sơ {new_ship['name']} thành công!")
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
                    save_all_ships_to_db()
                    st.rerun()

                if b2.button(
                    "🗑️",
                    key=f"d_{ship['id']}",
                    use_container_width=True
                ):
                    st.session_state.confirm_delete_id = ship["id"]
                    st.rerun()

            # Xác nhận trước khi xóa vĩnh viễn hồ sơ và chứng từ
            if st.session_state.get("confirm_delete_id") == ship["id"]:
                st.warning(
                    f"Xóa vĩnh viễn hồ sơ **{ship['name']}** "
                    f"(IMO {ship['imo']}) cùng toàn bộ chứng từ? "
                    "Thao tác này không thể hoàn tác."
                )
                y1, y2, _ = st.columns([1, 1, 4])

                if y1.button(
                    "Xóa",
                    type="primary",
                    key=f"d_yes_{ship['id']}",
                    use_container_width=True
                ):
                    folder = (
                        BASE_UPLOAD_DIR
                        / f"{ship['imo']}_{ship['id']}"
                    )

                    if folder.exists():
                        import shutil
                        shutil.rmtree(folder, ignore_errors=True)

                    # Xóa hồ sơ khỏi SQLite và danh sách phiên làm việc
                    delete_ship_from_db(ship["id"])

                    st.session_state.ships = [
                        s for s in st.session_state.ships
                        if s["id"] != ship["id"]
                    ]
                    st.session_state.confirm_delete_id = None
                    flash(f"Đã xóa hồ sơ {ship['name']}.")
                    st.rerun()

                if y2.button(
                    "Hủy",
                    key=f"d_no_{ship['id']}",
                    use_container_width=True
                ):
                    st.session_state.confirm_delete_id = None
                    st.rerun()

            st.divider()


# =========================================================
# 9. TRANG CHI TIẾT HỒ SƠ TÀU
# =========================================================

else:

    ship = get_ship(st.session_state.current_ship_id)

    if not ship:
        st.session_state.page = "list"
        save_all_ships_to_db()
        st.rerun()

    if st.button("← Danh sách hồ sơ"):
        st.session_state.page = "list"
        save_all_ships_to_db()
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
            save_all_ships_to_db()
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
            save_all_ships_to_db()
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
        c3.metric("Định dạng hỗ trợ", "PDF")

        st.divider()

        st.markdown("#### 📤 Tải lên chứng từ")

        # Đổi key sau mỗi lần tải để xóa danh sách file đã chọn
        uploader_round = st.session_state.get(f"upload_round_{ship['id']}", 0)

        uploaded_files = st.file_uploader(
            "Chọn một hoặc nhiều chứng từ",
            type=["pdf"],
            accept_multiple_files=True,
            key=f"upload_{ship['id']}_{uploader_round}"
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
                    if ship.get("status") == "Nháp":
                        ship["status"] = "Đã tải chứng từ"
                    flash(f"Đã tải thành công {success_count} chứng từ.")

                for error in errors:
                    flash(error, "warning")

                st.session_state[f"upload_round_{ship['id']}"] = (
                    uploader_round + 1
                )
                save_all_ships_to_db()
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

                file_path = resolve_doc_path(doc)
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

                            # Xóa luôn kết quả trích xuất và ánh xạ của
                            # chứng từ này để không xuất dữ liệu "mồ côi"
                            sgcl_ensure_ship_data(ship)
                            ship["field_extractions"] = [
                                rec for rec in ship["field_extractions"]
                                if rec.get("document") != doc["name"]
                            ]
                            ship["mappings"] = [
                                m for m in ship["mappings"]
                                if m.get("document") != doc["name"]
                                or m["id"] in ship["manual_entries"]
                            ]
                            sgcl_invalidate_review(ship)

                            if not ship["documents"]:
                                ship["status"] = "Nháp"
                                ship["step_status"]["2"] = "pending"
                                ship["step_status"]["3"] = "pending"

                            flash(f"Đã xóa chứng từ {doc['name']}.")
                            save_all_ships_to_db()
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

                    preview_path = resolve_doc_path(preview_doc)

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
                save_all_ships_to_db()
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
                save_all_ships_to_db()
                st.rerun()

    # =====================================================
    # BƯỚC 3 – TRÍCH XUẤT DỮ LIỆU TỪ PDF
    # =====================================================


    elif step == 3:
        sgcl_ensure_ship_data(ship)

        st.subheader("3. Trích xuất dữ liệu chứng từ")
        st.caption(
            "Chọn chứng từ để xem bản gốc và trích xuất các trường "
            "đã được cấu hình."
        )

        if not ship.get("documents"):
            st.warning("Hồ sơ chưa có chứng từ. Vui lòng quay lại Bước 2.")

        else:
            # ---------------------------------------------
            # CẤU HÌNH TRƯỜNG DỮ LIỆU
            # ---------------------------------------------
            with st.expander(
                "⚙️ Cấu hình trường dữ liệu dùng chung",
                expanded=False
            ):
                config = st.session_state.sgcl_field_config
                doc_types = list(config.keys())

                selected_config_type = st.selectbox(
                    "Loại chứng từ",
                    doc_types,
                    key=f"s3_config_type_{ship['id']}"
                )

                current_fields = config[selected_config_type]["fields"]

                config_df = pd.DataFrame([
                    {
                        "Tên trường": field["name"],
                        "Kiểu dữ liệu": field["type"]
                    }
                    for field in current_fields
                ])

                edited_config = st.data_editor(
                    config_df,
                    num_rows="dynamic",
                    use_container_width=True,
                    hide_index=True,
                    column_config={
                        "Kiểu dữ liệu": st.column_config.SelectboxColumn(
                            "Kiểu dữ liệu",
                            options=["single", "table"],
                            default="single",
                            required=True
                        )
                    },
                    key=f"s3_config_editor_{ship['id']}_{sgcl_normalize(selected_config_type)}"
                )

                # Key gắn với loại chứng từ, nếu không Streamlit giữ nguyên
                # từ khóa của loại cũ khi đổi loại -> lưu nhầm sang loại mới
                aliases_text = st.text_input(
                    "Từ khóa nhận diện tên file, cách nhau bằng dấu phẩy",
                    value=", ".join(
                        config[selected_config_type].get("aliases", [])
                    ),
                    key=f"s3_aliases_{ship['id']}_{sgcl_normalize(selected_config_type)}"
                )

                if st.button(
                    "Lưu cấu hình",
                    type="primary",
                    key=f"s3_save_config_{ship['id']}"
                ):
                    new_fields = []

                    for _, row in edited_config.fillna("").iterrows():
                        field_name = str(row["Tên trường"]).strip()
                        field_type = str(row["Kiểu dữ liệu"]).strip().lower()

                        if not field_name:
                            continue

                        if field_type not in ["single", "table"]:
                            field_type = "single"

                        new_fields.append({
                            "name": field_name,
                            "type": field_type
                        })

                    config[selected_config_type]["fields"] = new_fields
                    config[selected_config_type]["aliases"] = [
                        x.strip()
                        for x in aliases_text.split(",")
                        if x.strip()
                    ]

                    sgcl_save_field_config(config)
                    st.session_state.sgcl_field_config = config

                    flash("Đã lưu cấu hình trường dữ liệu.")
                    st.rerun()

                st.caption(
                    "single = một giá trị; table = dữ liệu nhiều dòng."
                )

            st.divider()

            # ---------------------------------------------
            # CHỌN CHỨNG TỪ
            # ---------------------------------------------
            documents = ship["documents"]
            document_names = [doc["name"] for doc in documents]

            selected_name = st.selectbox(
                "Chọn chứng từ",
                document_names,
                key=f"s3_document_{ship['id']}"
            )

            selected_doc = next(
                doc for doc in documents
                if doc["name"] == selected_name
            )

            config = st.session_state.sgcl_field_config

            detected_type = sgcl_get_document_type(
                selected_name,
                config
            )

            type_options = list(config.keys())

            default_index = (
                type_options.index(detected_type)
                if detected_type in type_options else 0
            )

            chosen_type = st.selectbox(
                "Loại chứng từ",
                type_options,
                index=default_index,
                key=f"s3_type_{ship['id']}_{sgcl_normalize(selected_name)}"
            )

            if detected_type:
                st.caption(
                    f"Hệ thống nhận diện tên file: {detected_type}. "
                    "Bà có thể chọn lại nếu chưa chính xác."
                )
            else:
                st.warning(
                    "Không nhận diện được loại chứng từ từ tên file. "
                    "Vui lòng chọn thủ công."
                )

            fields = config[chosen_type]["fields"]

            file_path = resolve_doc_path(selected_doc)

            if not file_path.exists():
                st.error("Không tìm thấy file PDF trong thư mục hồ sơ.")
            else:
                file_bytes = file_path.read_bytes()

                # -----------------------------------------
                # HIỂN THỊ PDF VÀ TRƯỜNG CẦN TRÍCH XUẤT
                # -----------------------------------------
                left, right = st.columns([1.05, 1])

                with left:
                    st.markdown("#### 📄 Chứng từ gốc")

                    with fitz.open(
                        stream=file_bytes,
                        filetype="pdf"
                    ) as pdf:
                        total_pages = len(pdf)

                    page_number = st.number_input(
                        "Trang PDF",
                        min_value=1,
                        max_value=max(1, total_pages),
                        value=1,
                        key=f"s3_preview_page_{ship['id']}_{sgcl_normalize(selected_name)}"
                    )

                    with fitz.open(
                        stream=file_bytes,
                        filetype="pdf"
                    ) as pdf:
                        page = pdf[int(page_number) - 1]
                        z = sgcl_safe_zoom(page, 1.5, max_side=1800)
                        pix = page.get_pixmap(
                            matrix=fitz.Matrix(z, z),
                            alpha=False
                        )
                        st.image(
                            pix.tobytes("png"),
                            use_container_width=True
                        )

                    st.download_button(
                        "⬇️ Tải chứng từ gốc",
                        data=file_bytes,
                        file_name=selected_name,
                        mime="application/pdf",
                        key=f"s3_download_{ship['id']}_{sgcl_normalize(selected_name)}"
                    )

                with right:
                    st.markdown("#### 🧾 Trường dữ liệu")

                    if not fields:
                        st.warning(
                            "Chưa có trường dữ liệu cho loại chứng từ này."
                        )

                    else:
                        for field in fields:
                            st.write(
                                ("▦ " if field["type"] == "table" else "• ")
                                + field["name"]
                            )

                        if st.button(
                            "🔍 Trích xuất toàn bộ trường",
                            type="primary",
                            use_container_width=True,
                            key=f"s3_extract_{ship['id']}_{sgcl_normalize(selected_name)}"
                        ):
                            with st.spinner(
                                "Đang đọc toàn bộ trang của chứng từ..."
                            ):
                                try:
                                    pages = sgcl_get_all_pdf_text(
                                        file_bytes
                                    )

                                    # Xóa kết quả cũ của đúng chứng từ
                                    # và đúng loại chứng từ
                                    ship["field_extractions"] = [
                                        rec for rec in ship["field_extractions"]
                                        if not (
                                            rec.get("document") == selected_name
                                            and rec.get("doc_type") == chosen_type
                                        )
                                    ]

                                    single_fields = [
                                        field for field in fields
                                        if field["type"] == "single"
                                    ]

                                    table_fields = [
                                        field for field in fields
                                        if field["type"] == "table"
                                    ]

                                    # Trường đơn
                                    for field in single_fields:
                                        result = sgcl_extract_single(
                                            pages,
                                            field["name"],
                                            other_fields=[f["name"] for f in fields]
                                        )

                                        ship["field_extractions"].append({
                                            "id": str(uuid.uuid4()),
                                            "document": selected_name,
                                            "doc_type": chosen_type,
                                            "field": field["name"],
                                            "type": "single",
                                            "result": {
                                                "value": result["value"]
                                            },
                                            "page": result["page"],
                                            "found": result["found"],
                                            "source_type": "extracted"
                                        })

                                    # Nhóm các trường bảng thành một bảng
                                    if table_fields:
                                        headers = [
                                            field["name"]
                                            for field in table_fields
                                        ]

                                        table_result = sgcl_extract_table(
                                            pages,
                                            headers,
                                            file_bytes=file_bytes
                                        )

                                        ship["field_extractions"].append({
                                            "id": str(uuid.uuid4()),
                                            "document": selected_name,
                                            "doc_type": chosen_type,
                                            "field": f"Bảng dữ liệu - {chosen_type}",
                                            "type": "table",
                                            "result": table_result,
                                            "page": table_result["page"],
                                            "found": table_result["found"],
                                            "source_type": "extracted"
                                        })

                                    sgcl_invalidate_review(ship)
                                    ship["step_status"]["3"] = "done"

                                    if ship.get("status") != "Hoàn thành":
                                        ship["status"] = "Đã trích xuất"

                                    found_count = sum(
                                        1 for rec in ship["field_extractions"]
                                        if rec.get("document") == selected_name
                                        and rec.get("found")
                                    )
                                    total_count = sum(
                                        1 for rec in ship["field_extractions"]
                                        if rec.get("document") == selected_name
                                    )
                                    flash(
                                        f"Đã trích xuất xong: tìm thấy "
                                        f"{found_count}/{total_count} mục. "
                                        "Bà kiểm tra và chỉnh sửa kết quả bên dưới."
                                    )
                                    save_all_ships_to_db()
                                    st.rerun()

                                except Exception as e:
                                    st.error(
                                        f"Lỗi khi trích xuất dữ liệu: {e}"
                                    )

                st.divider()

                # -----------------------------------------
                # HIỂN THỊ VÀ CHỈNH SỬA KẾT QUẢ
                # -----------------------------------------
                st.markdown("#### Kết quả trích xuất")

                records = [
                    rec for rec in ship["field_extractions"]
                    if rec.get("document") == selected_name
                ]

                if not records:
                    st.info(
                        "Chưa có kết quả. Nhấn nút Trích xuất toàn bộ trường."
                    )

                for rec in records:
                    with st.container(border=True):
                        a, b = st.columns([3, 1])

                        a.markdown(f"**{rec['field']}**")
                        a.caption(
                            f"Nguồn: {rec.get('document', '')} "
                            f"· Trang: {rec.get('page') or 'Chưa xác định'}"
                        )

                        if rec.get("found"):
                            b.success("Đã tìm thấy")
                        else:
                            b.warning("Chưa tìm thấy")

                        if rec["type"] == "single":
                            old_value = rec["result"].get("value", "")

                            new_value = st.text_area(
                                "Giá trị trích xuất",
                                value=str(old_value),
                                height=75,
                                key=f"s3_edit_{ship['id']}_{rec['id']}"
                            )

                            if new_value != old_value:
                                rec["result"]["value"] = new_value
                                rec["source_type"] = "manually_edited"
                                rec["found"] = bool(new_value.strip())
                                sgcl_invalidate_review(ship)

                        else:
                            table = rec["result"]
                            headers = table.get("headers", [])
                            rows = table.get("rows", [])

                            if not headers:
                                headers = ["Thông tin"]
                                table["headers"] = headers

                            st.caption(f"{len(rows)} dòng")

                            new_rows, changed = sgcl_table_editor(
                                f"s3_table_{ship['id']}_{rec['id']}",
                                headers,
                                rows
                            )

                            if changed:
                                rec["result"]["rows"] = new_rows
                                rec["source_type"] = "manually_edited"
                                rec["found"] = bool(new_rows)
                                sgcl_invalidate_review(ship)

                        if st.button(
                            "🗑️ Xóa kết quả này",
                            key=f"s3_delete_{ship['id']}_{rec['id']}"
                        ):
                            ship["field_extractions"] = [
                                x for x in ship["field_extractions"]
                                if x["id"] != rec["id"]
                            ]

                            sgcl_invalidate_review(ship)
                            save_all_ships_to_db()
                            st.rerun()

        st.divider()

        col1, col2 = st.columns(2)

        if col1.button(
            "← Quay lại Bước 2",
            use_container_width=True
        ):
            ship["current_step"] = 2
            save_all_ships_to_db()
            st.rerun()

        if col2.button(
            "Xác nhận & Tiếp tục →",
            type="primary",
            use_container_width=True,
            disabled=not ship.get("field_extractions")
        ):
            ship["step_status"]["3"] = "done"
            ship["current_step"] = 4
            save_all_ships_to_db()
            st.rerun()

    # =====================================================
    # BƯỚC 4 – KIỂM TRA VÀ SẮP XẾP DỮ LIỆU
    # =====================================================

    elif step == 4:
        sgcl_ensure_ship_data(ship)

        st.subheader("4. Kiểm tra, đối chiếu và ánh xạ Excel")
        st.caption(
            "Kiểm tra dữ liệu chứng từ, chỉnh sửa giá trị cuối cùng, "
            "nhập bổ sung và chọn vị trí ghi vào biểu mẫu."
        )

        # ---------------------------------------------
        # TẠO MAPPING TỪ KẾT QUẢ BƯỚC 3
        # ---------------------------------------------
        import copy

        extraction_lookup = {
            rec["id"]: rec
            for rec in ship["field_extractions"]
        }

        # Mapping của kết quả đã bị xóa/trích xuất lại ở Bước 3 -> gỡ bỏ,
        # nhưng giữ lại vị trí sheet/ô để áp cho kết quả mới cùng trường.
        old_positions = {}
        kept_mappings = []

        for m in ship["mappings"]:
            source_id = m.get("source_record_id")
            if source_id and source_id not in extraction_lookup:
                old_positions[(m.get("document"), m.get("field"))] = m
                continue
            kept_mappings.append(m)

        if len(kept_mappings) != len(ship["mappings"]):
            ship["mappings"] = kept_mappings
            sgcl_invalidate_review(ship)

        existing_ids = {
            item.get("source_record_id")
            for item in ship["mappings"]
        }

        mapping_memory = sgcl_load_mapping_memory()

        for rec in ship["field_extractions"]:
            if rec["id"] in existing_ids:
                continue

            old = old_positions.get((rec.get("document", ""), rec["field"]))
            remembered = mapping_memory.get(
                sgcl_mapping_key(rec.get("doc_type", ""), rec["field"]),
                {}
            )
            position = old or remembered

            is_single = rec["type"] == "single"
            source_value = (
                rec.get("result", {}).get("value", "") if is_single else ""
            )
            source_table = (
                None if is_single
                else copy.deepcopy(rec.get("result", {}))
            )

            ship["mappings"].append({
                "id": str(uuid.uuid4()),
                "source_record_id": rec["id"],
                "field": rec["field"],
                "type": rec["type"],
                "doc_type": rec.get("doc_type", ""),
                "document": rec.get("document", ""),
                "source_type": rec.get("source_type", "extracted"),
                "value": source_value,
                "table": source_table,
                # Ảnh chụp dữ liệu nguồn để biết Bước 3 có sửa lại hay không
                "source_snapshot": (
                    source_value if is_single
                    else copy.deepcopy(source_table.get("rows", []))
                ),
                "sheet": (
                    position.get("sheet")
                    if position.get("sheet") in NSW_SHEETS
                    else NSW_SHEETS[0]
                ),
                # Để trống = chưa ánh xạ, tránh mặc định A1 ghi đè tiêu đề mẫu
                "cell": position.get("cell", ""),
                "extra_targets": (
                    list(position.get("extra_targets", []) or [])
                    if is_single else []
                ),
                "column_map": dict(position.get("column_map", {}) or {})
            })

        # Đồng bộ chỉnh sửa từ Bước 3 — CHỈ khi dữ liệu nguồn thực sự đổi,
        # để không ghi đè "Giá trị cuối cùng" đã sửa ở Bước 4.
        for mapping in ship["mappings"]:
            source_id = mapping.get("source_record_id")

            if not source_id or source_id not in extraction_lookup:
                continue

            rec = extraction_lookup[source_id]
            mapping["field"] = rec["field"]
            mapping["type"] = rec["type"]
            mapping["document"] = rec.get("document", "")
            mapping.setdefault("doc_type", rec.get("doc_type", ""))

            if rec["type"] == "single":
                source_value = rec["result"].get("value", "")
                if source_value != mapping.get("source_snapshot"):
                    mapping["value"] = source_value
                    mapping["source_snapshot"] = source_value
                    mapping["source_type"] = rec.get("source_type", "extracted")
            else:
                source_rows = rec.get("result", {}).get("rows", [])
                if source_rows != mapping.get("source_snapshot"):
                    mapping["table"] = copy.deepcopy(rec.get("result", {}))
                    mapping["source_snapshot"] = copy.deepcopy(source_rows)
                    mapping["source_type"] = rec.get("source_type", "extracted")

        # ---------------------------------------------
        # ĐỐI CHIẾU GIÁ TRỊ TRÙNG TÊN TRƯỜNG
        # ---------------------------------------------
        grouped_values = {}

        for rec in ship["field_extractions"]:
            if rec["type"] != "single":
                continue

            value = str(
                rec.get("result", {}).get("value", "")
            ).strip()

            if value:
                key = sgcl_normalize(rec["field"])
                grouped_values.setdefault(key, []).append({
                    "field": rec["field"],
                    "value": value,
                    "document": rec.get("document", "")
                })

        conflicts = []

        for key, items in grouped_values.items():
            unique_values = {
                item["value"].casefold()
                for item in items
            }

            if len(unique_values) > 1:
                conflicts.append(items)

        with st.expander(
            f"🔎 Đối chiếu giữa các chứng từ - {len(conflicts)} trường cần kiểm tra",
            expanded=bool(conflicts)
        ):
            if not conflicts:
                st.success(
                    "Chưa phát hiện trường đơn có giá trị khác nhau."
                )

            for items in conflicts:
                st.warning(items[0]["field"])

                for item in items:
                    st.write(
                        f"• **{item['value']}** — {item['document']}"
                    )

                st.caption(
                    "Hệ thống không tự chọn giá trị. "
                    "Bà quyết định giá trị cuối cùng ở phần ánh xạ bên dưới."
                )

        st.divider()

        # ---------------------------------------------
        # DANH SÁCH DỮ LIỆU VÀ ÁNH XẠ EXCEL
        # ---------------------------------------------
        st.markdown("#### Dữ liệu và vị trí ghi Excel")
        st.caption(
            "Để trống ô đích nếu không muốn xuất trường đó. "
            "Vị trí đã chọn được ghi nhớ cho các hồ sơ tàu sau."
        )

        if not ship["mappings"]:
            st.info(
                "Chưa có dữ liệu. Bà có thể nhập thủ công ở phần bên dưới."
            )

        # Áp vị trí mặc định/ghi nhớ cho các trường chưa có ô (hồ sơ cũ)
        not_mapped = [
            m for m in ship["mappings"]
            if not str(m.get("cell", "")).strip()
            or str(m.get("cell", "")).upper() == "A1"
        ]
        if not_mapped:
            memory_for_fill = sgcl_load_mapping_memory()
            fillable = [
                m for m in not_mapped
                if sgcl_mapping_key(m.get("doc_type", ""), m["field"])
                in memory_for_fill
            ]
            if fillable and st.button(
                f"📌 Điền vị trí theo TemplateNSW cho {len(fillable)} "
                "trường chưa chọn ô",
                key=f"s4_fill_default_{ship['id']}"
            ):
                for m in fillable:
                    pos = memory_for_fill[
                        sgcl_mapping_key(m.get("doc_type", ""), m["field"])
                    ]
                    m["sheet"] = pos.get("sheet", NSW_SHEETS[0])
                    m["cell"] = pos.get("cell", "")
                    m["column_map"] = dict(pos.get("column_map", {}) or {})
                    if m["type"] == "single":
                        m["extra_targets"] = list(
                            pos.get("extra_targets", []) or []
                        )
                    # Xóa trạng thái ô nhập cũ để hiển thị vị trí mới
                    for k in list(st.session_state.keys()):
                        if m["id"] in str(k) and str(k).startswith(
                            ("s4_sheet_", "s4_cell_", "s4_extra_", "s4_column_")
                        ):
                            del st.session_state[k]
                sgcl_invalidate_review(ship)
                flash(f"Đã điền vị trí cho {len(fillable)} trường.")
                save_all_ships_to_db()
                st.rerun()

        # Xem trực tiếp lưới ô của biểu mẫu gốc và chọn ô bằng thao tác click.
        # st.dataframe hỗ trợ chọn một ô trên các phiên bản Streamlit mới.
        if TEMPLATE_PATH.exists() and ship["mappings"]:
            st.markdown("#### Xem trước biểu mẫu Excel gốc – chọn ô đích")
            st.caption("Chọn trường cần ánh xạ, chọn sheet, rồi click ô tương ứng trong lưới bên dưới. Địa chỉ ô sẽ tự điền vào trường đó.")
            mapping_options = {
                f"{m['field']} · {m.get('document') or 'Nhập thủ công'}": m["id"]
                for m in ship["mappings"]
            }
            selected_mapping_label = st.selectbox(
                "Trường đang chọn để ánh xạ",
                list(mapping_options.keys()),
                key=f"s4_preview_mapping_{ship['id']}"
            )
            selected_mapping_id = mapping_options[selected_mapping_label]
            selected_mapping_obj = next(m for m in ship["mappings"] if m["id"] == selected_mapping_id)
            preview_sheet = st.selectbox(
                "Sheet xem trước",
                NSW_SHEETS,
                index=(NSW_SHEETS.index(selected_mapping_obj["sheet"]) if selected_mapping_obj.get("sheet") in NSW_SHEETS else 0),
                key=f"s4_preview_sheet_{ship['id']}"
            )
            try:
                preview_wb = openpyxl.load_workbook(TEMPLATE_PATH, read_only=True, data_only=False)
                preview_ws = preview_wb[preview_sheet]
                max_row = min(max(preview_ws.max_row or 1, 1), 60)
                max_col = min(max(preview_ws.max_column or 1, 1), 24)
                preview_data = []
                for row_num in range(1, max_row + 1):
                    preview_row = []
                    for col_num in range(1, max_col + 1):
                        value = preview_ws.cell(row_num, col_num).value
                        preview_row.append("" if value is None else str(value)[:120])
                    preview_data.append(preview_row)
                preview_wb.close()
                preview_df = pd.DataFrame(
                    preview_data,
                    index=[str(i) for i in range(1, max_row + 1)],
                    columns=[openpyxl.utils.get_column_letter(i) for i in range(1, max_col + 1)]
                )
                preview_event = st.dataframe(
                    preview_df, use_container_width=True, height=420,
                    on_select="rerun", selection_mode="single-cell",
                    key=f"s4_excel_grid_{ship['id']}"
                )
                selected_cells = getattr(getattr(preview_event, "selection", None), "cells", []) or []
                if selected_cells:
                    picked = selected_cells[-1]
                    row_idx = picked.get("row")
                    col_idx = picked.get("column")
                    if isinstance(row_idx, int) and isinstance(col_idx, int):
                        cell_address = f"{openpyxl.utils.get_column_letter(col_idx + 1)}{row_idx + 1}"
                        if (selected_mapping_obj.get("sheet") != preview_sheet or selected_mapping_obj.get("cell") != cell_address):
                            selected_mapping_obj["sheet"] = preview_sheet
                            selected_mapping_obj["cell"] = cell_address
                            sgcl_invalidate_review(ship)
                            for state_key in (f"s4_sheet_{ship['id']}_{selected_mapping_id}", f"s4_cell_{ship['id']}_{selected_mapping_id}"):
                                st.session_state.pop(state_key, None)
                            flash(f"Đã chọn ô {preview_sheet}!{cell_address} cho trường {selected_mapping_obj['field']}.")
                            save_all_ships_to_db()
                            st.rerun()
                st.caption("Lưới hiển thị tối đa 60 hàng × 24 cột đầu tiên để thao tác thuận tiện; file xuất vẫn giữ nguyên cấu trúc workbook gốc.")
            except Exception as preview_error:
                st.warning(f"Chưa thể hiển thị lưới Excel tương tác: {preview_error}. Có thể nhập địa chỉ ô ở phần ánh xạ bên dưới.")

        for mapping in ship["mappings"]:
            with st.container(border=True):
                left, right = st.columns([1, 1])

                with left:
                    st.markdown(f"**{mapping['field']}**")

                    st.caption(
                        f"Chứng từ nguồn: "
                        f"{mapping.get('document') or 'Nhập thủ công'}"
                    )

                    source = mapping.get("source_type", "extracted")

                    if source == "extracted":
                        st.info("Dữ liệu trích xuất")
                    elif source == "manually_edited":
                        st.warning("Dữ liệu đã chỉnh sửa")
                    else:
                        st.info("Dữ liệu nhập thủ công")

                    if mapping["type"] == "single":
                        old_value = str(mapping.get("value", ""))

                        final_value = st.text_area(
                            "Giá trị cuối cùng",
                            value=old_value,
                            height=80,
                            key=f"s4_value_{ship['id']}_{mapping['id']}"
                        )

                        if final_value != old_value:
                            mapping["value"] = final_value

                            if mapping["source_type"] == "extracted":
                                mapping["source_type"] = "manually_edited"

                            sgcl_invalidate_review(ship)

                    else:
                        table = mapping.get("table") or {}
                        headers = table.get("headers", []) or ["Thông tin"]
                        table["headers"] = headers
                        mapping["table"] = table

                        st.caption(f"{len(table.get('rows', []))} dòng")

                        new_rows, changed = sgcl_table_editor(
                            f"s4_table_edit_{ship['id']}_{mapping['id']}",
                            headers,
                            table.get("rows", [])
                        )

                        if changed:
                            table["rows"] = new_rows
                            if mapping["source_type"] == "extracted":
                                mapping["source_type"] = "manually_edited"
                            sgcl_invalidate_review(ship)

                with right:
                    st.markdown("**Ánh xạ vào biểu mẫu NSW**")

                    sheet_index = (
                        NSW_SHEETS.index(mapping["sheet"])
                        if mapping.get("sheet") in NSW_SHEETS else 0
                    )

                    selected_sheet = st.selectbox(
                        "Sheet đích",
                        NSW_SHEETS,
                        index=sheet_index,
                        key=f"s4_sheet_{ship['id']}_{mapping['id']}"
                    )

                    selected_cell = st.text_input(
                        "Ô đích / ô bắt đầu (VD: B5)",
                        value=mapping.get("cell", ""),
                        key=f"s4_cell_{ship['id']}_{mapping['id']}"
                    ).strip().upper()

                    cell_valid = True

                    if selected_cell:
                        try:
                            openpyxl.utils.cell.coordinate_from_string(
                                selected_cell
                            )
                        except Exception:
                            cell_valid = False
                            st.error("Địa chỉ ô không hợp lệ. Ví dụ: B5, C10.")
                    else:
                        st.caption("⚪ Chưa chọn ô — trường này sẽ không được xuất.")

                    # Đổi vị trí ghi sau khi đã xác nhận -> phải xác nhận lại
                    if (
                        selected_sheet != mapping.get("sheet")
                        or selected_cell != mapping.get("cell", "")
                    ):
                        sgcl_invalidate_review(ship)

                    mapping["sheet"] = selected_sheet
                    mapping["cell"] = selected_cell

                    if mapping["type"] == "single":
                        old_extra = sgcl_format_targets(
                            mapping.get("extra_targets", [])
                        )
                        extra_text = st.text_input(
                            "Ghi thêm vào (tùy chọn)",
                            value=old_extra,
                            placeholder="VD: Notice ship to port!B24; ConfirmationOfArrival!B6",
                            help=(
                                "Cùng một dữ liệu có thể cần điền ở nhiều bản "
                                "khai. Dạng Sheet!Ô, ngăn cách bởi dấu ;"
                            ),
                            key=f"s4_extra_{ship['id']}_{mapping['id']}"
                        )
                        extra_targets, extra_errors = sgcl_parse_targets(
                            extra_text
                        )
                        for err in extra_errors:
                            st.error(err)
                        if extra_errors:
                            cell_valid = False

                        if extra_text.strip() != old_extra.strip():
                            sgcl_invalidate_review(ship)
                        if not extra_errors:
                            mapping["extra_targets"] = extra_targets

                    if mapping["type"] == "table":
                        headers = (mapping.get("table") or {}).get("headers", [])

                        st.caption(
                            "Cột Excel cho từng trường bảng. Dòng đầu tiên "
                            "của bảng được ghi vào hàng của ô bắt đầu."
                        )

                        try:
                            start_col_letter, _ = (
                                openpyxl.utils.cell.coordinate_from_string(
                                    selected_cell
                                )
                            )
                            start_col_number = (
                                openpyxl.utils.column_index_from_string(
                                    start_col_letter
                                )
                            )
                        except Exception:
                            start_col_number = 1

                        column_map = mapping.get("column_map", {}) or {}

                        for index, header in enumerate(headers):
                            default_column = (
                                openpyxl.utils.get_column_letter(
                                    start_col_number + index
                                )
                            )

                            column_value = st.text_input(
                                f"Cột Excel cho: {header}",
                                value=column_map.get(
                                    header,
                                    default_column
                                ),
                                key=f"s4_column_{ship['id']}_{mapping['id']}_{index}"
                            ).strip().upper()

                            if not re.fullmatch(r"[A-Z]{1,3}", column_value):
                                cell_valid = False
                                st.error(
                                    f"Cột '{column_value}' không hợp lệ. "
                                    "Chỉ nhập chữ cái, VD: B, AA."
                                )

                            if column_map.get(header) != column_value:
                                sgcl_invalidate_review(ship)

                            column_map[header] = column_value

                        mapping["column_map"] = column_map

                    mapping["_invalid"] = not cell_valid

                # Xóa mapping thủ công
                if mapping["id"] in ship["manual_entries"]:
                    if st.button(
                        "🗑️ Xóa dữ liệu thủ công",
                        key=f"s4_delete_manual_{ship['id']}_{mapping['id']}"
                    ):
                        ship["mappings"] = [
                            x for x in ship["mappings"]
                            if x["id"] != mapping["id"]
                        ]

                        ship["manual_entries"].remove(mapping["id"])
                        sgcl_invalidate_review(ship)
                        save_all_ships_to_db()
                        st.rerun()

        # ---------------------------------------------
        # THÊM DỮ LIỆU KHÔNG CÓ TRONG CHỨNG TỪ
        # ---------------------------------------------
        st.divider()
        st.markdown("#### ➕ Nhập dữ liệu thủ công")

        with st.form(
            f"s4_manual_form_{ship['id']}",
            clear_on_submit=True
        ):
            manual_name = st.text_input(
                "Tên trường dữ liệu *"
            )

            manual_type = st.radio(
                "Loại dữ liệu",
                ["Một giá trị", "Bảng nhiều dòng"],
                horizontal=True
            )

            manual_value = st.text_area(
                "Giá trị",
                height=80,
                help=(
                    "Với bảng nhiều dòng: mỗi dòng là một hàng; "
                    "dòng đầu là tiêu đề cột; các cột ngăn bởi dấu | "
                    "hoặc Tab (có thể dán trực tiếp từ Excel)."
                )
            )

            manual_sheet = st.selectbox(
                "Sheet Excel",
                NSW_SHEETS,
                key=f"s4_manual_sheet_{ship['id']}"
            )

            manual_cell = st.text_input(
                "Ô đích / ô bắt đầu",
                placeholder="VD: B5"
            )

            add_manual = st.form_submit_button(
                "➕ Thêm dữ liệu",
                type="primary"
            )

            if add_manual:
                cell_text = manual_cell.strip().upper()

                if not manual_name.strip():
                    st.error("Vui lòng nhập tên trường.")
                else:
                    try:
                        openpyxl.utils.cell.coordinate_from_string(cell_text)
                        cell_ok = True
                    except Exception:
                        cell_ok = False

                    if not cell_ok:
                        st.error(
                            "Ô Excel không hợp lệ. Ví dụ: A1, B5, C10."
                        )
                    else:
                        is_table = manual_type == "Bảng nhiều dòng"
                        table_data = None

                        if is_table:
                            lines = [
                                line for line in manual_value.splitlines()
                                if line.strip()
                            ]
                            parsed = [
                                [c.strip() for c in re.split(r"\t|\|", line)]
                                for line in lines
                            ]
                            if parsed:
                                headers = parsed[0]
                                table_data = {
                                    "headers": headers,
                                    "rows": sgcl_normalize_rows(
                                        parsed[1:], headers
                                    )
                                }
                            else:
                                table_data = {
                                    "headers": [manual_name.strip()],
                                    "rows": []
                                }

                        new_mapping = {
                            "id": str(uuid.uuid4()),
                            "source_record_id": None,
                            "field": manual_name.strip(),
                            "type": "table" if is_table else "single",
                            "doc_type": "Nhập thủ công",
                            "document": "",
                            "source_type": "manually_entered",
                            "value": "" if is_table else manual_value,
                            "table": table_data,
                            "sheet": manual_sheet,
                            "cell": cell_text,
                            "column_map": {}
                        }

                        ship["mappings"].append(new_mapping)
                        ship["manual_entries"].append(
                            new_mapping["id"]
                        )

                        sgcl_invalidate_review(ship)
                        flash("Đã thêm dữ liệu thủ công.")
                        save_all_ships_to_db()
                        st.rerun()

        # ---------------------------------------------
        # KIỂM TRA TRƯỚC KHI XÁC NHẬN
        # ---------------------------------------------
        st.divider()
        st.markdown("#### Kiểm tra trước khi xác nhận")

        invalid_mappings = [
            m for m in ship["mappings"]
            if m.get("_invalid")
        ]

        unmapped = [
            m for m in ship["mappings"]
            if not str(m.get("cell", "")).strip()
        ]

        missing_values = [
            m for m in ship["mappings"]
            if m["type"] == "single"
            and str(m.get("cell", "")).strip()
            and not str(m.get("value", "")).strip()
        ]

        # Hai trường khác nhau cùng ghi vào một ô -> trường sau ghi đè
        used_cells = {}
        for m in ship["mappings"]:
            if m["type"] == "single" and str(m.get("cell", "")).strip():
                for target in (
                    [{"sheet": m["sheet"], "cell": m["cell"]}]
                    + list(m.get("extra_targets", []) or [])
                ):
                    used_cells.setdefault(
                        (target["sheet"], target["cell"]), []
                    ).append(m["field"])

        duplicate_cells = {
            k: v for k, v in used_cells.items() if len(v) > 1
        }

        if invalid_mappings:
            st.error(
                f"Có {len(invalid_mappings)} địa chỉ ô/cột Excel không hợp lệ."
            )

        if duplicate_cells:
            for (sheet_name, cell), fields_in_cell in duplicate_cells.items():
                st.error(
                    f"Ô {cell} (sheet '{sheet_name}') được gán cho nhiều "
                    f"trường: {', '.join(fields_in_cell)}."
                )

        if unmapped:
            st.info(
                f"Có {len(unmapped)} trường chưa chọn ô Excel và sẽ "
                "không được xuất."
            )

        if missing_values:
            st.warning(
                f"Có {len(missing_values)} trường đang để trống. "
                "Các trường này sẽ không được tự động điền số 0."
            )

        if conflicts:
            st.warning(
                f"Có {len(conflicts)} trường cần kiểm tra do dữ liệu "
                "giữa các chứng từ không giống nhau."
            )

        st.info(
            "Dữ liệu trích xuất, chỉnh sửa và nhập thủ công được "
            "phân biệt riêng. Hệ thống không tự suy đoán dữ liệu còn thiếu."
        )

        if st.button(
            "Xác nhận dữ liệu & Tiếp tục →",
            type="primary",
            use_container_width=True,
            disabled=(
                len(unmapped) == len(ship["mappings"])
                or bool(invalid_mappings)
                or bool(duplicate_cells)
            )
        ):
            ship["review_confirmed"] = True
            ship["step_status"]["4"] = "done"
            ship["status"] = "Đã xác nhận"
            ship["current_step"] = 5

            try:
                sgcl_save_mapping_memory(ship)
            except OSError:
                pass

            save_all_ships_to_db()
            st.rerun()

        st.divider()

        if st.button(
            "← Quay lại Bước 3",
            use_container_width=True,
            key=f"s4_back_{ship['id']}"
        ):
            ship["current_step"] = 3
            save_all_ships_to_db()
            st.rerun()

    # =====================================================
    # BƯỚC 5 – XUẤT FILE EXCEL NSW
    # =====================================================


    elif step == 5:
        sgcl_ensure_ship_data(ship)

        st.subheader("5. Xuất file Excel NSW")
        st.caption(
            "Tạo file Excel riêng cho hồ sơ tàu từ biểu mẫu TemplateNSW.xlsx."
        )

        if not ship.get("review_confirmed"):
            st.warning(
                "Bà cần hoàn tất và xác nhận dữ liệu tại Bước 4 trước."
            )

        elif not TEMPLATE_PATH.exists():
            st.error(
                "Không tìm thấy TemplateNSW.xlsx. "
                "Hãy đặt file mẫu cùng thư mục với app.py."
            )

        else:
            st.success(
                "Đã tìm thấy biểu mẫu TemplateNSW.xlsx."
            )

            st.markdown("#### Tóm tắt dữ liệu chuẩn bị xuất")

            summary = []

            for mapping in ship["mappings"]:
                summary.append({
                    "Trường dữ liệu": mapping["field"],
                    "Sheet": mapping.get("sheet", ""),
                    "Ô bắt đầu": mapping.get("cell", "") or "— không xuất —",
                    "Nguồn": mapping.get("source_type", ""),
                    "Kiểu dữ liệu": (
                        "Bảng nhiều dòng"
                        if mapping["type"] == "table"
                        else "Một giá trị"
                    )
                })

            if summary:
                st.dataframe(
                    pd.DataFrame(summary),
                    use_container_width=True,
                    hide_index=True
                )

            # -----------------------------------------
            # CẢNH BÁO DỮ LIỆU THIẾU
            # -----------------------------------------
            missing_values = [
                m for m in ship["mappings"]
                if m["type"] == "single"
                and str(m.get("cell", "")).strip()
                and not str(m.get("value", "")).strip()
            ]

            if missing_values:
                with st.expander(
                    f"⚠️ Có {len(missing_values)} trường đang để trống",
                    expanded=True
                ):
                    for item in missing_values:
                        st.write(
                            f"• {item['field']} — {item.get('document') or 'Nhập thủ công'}"
                        )

                    st.caption(
                        "Những ô có giá trị trống sẽ được giữ nguyên trong "
                        "biểu mẫu, không tự điền số 0."
                    )

            # -----------------------------------------
            # XUẤT EXCEL
            # -----------------------------------------
            if st.button(
                "📊 Tạo file Excel NSW",
                type="primary",
                use_container_width=True,
                key=f"s5_create_{ship['id']}"
            ):
                try:
                    wb = openpyxl.load_workbook(TEMPLATE_PATH)

                    # Kiểm tra đủ 10 Sheet được phép sử dụng
                    missing_sheets = [
                        name for name in NSW_SHEETS
                        if name not in wb.sheetnames
                    ]

                    if missing_sheets:
                        raise ValueError(
                            "Template thiếu Sheet: "
                            + ", ".join(missing_sheets)
                        )

                    skipped_fields = []
                    written_count = 0

                    # Ghi dữ liệu vào workbook đã mở một lần
                    for mapping in ship["mappings"]:
                        sheet_name = mapping.get("sheet")

                        if sheet_name not in NSW_SHEETS:
                            raise ValueError(
                                f"Sheet không được phép ghi: {sheet_name}"
                            )

                        if sheet_name not in wb.sheetnames:
                            raise ValueError(
                                f"Không tìm thấy Sheet: {sheet_name}"
                            )

                        ws = wb[sheet_name]

                        start_cell = str(
                            mapping.get("cell", "")
                        ).strip().upper()

                        # Trường chưa chọn ô -> không xuất
                        if not start_cell:
                            skipped_fields.append(mapping["field"])
                            continue

                        try:
                            start_col_letter, start_row = (
                                openpyxl.utils.cell.coordinate_from_string(
                                    start_cell
                                )
                            )
                        except Exception:
                            raise ValueError(
                                f"Địa chỉ ô không hợp lệ: {start_cell}"
                            )

                        start_col = (
                            openpyxl.utils.column_index_from_string(
                                start_col_letter
                            )
                        )

                        # ---------------------------------
                        # DỮ LIỆU MỘT GIÁ TRỊ
                        # ---------------------------------
                        if mapping["type"] == "single":
                            value = str(mapping.get("value", ""))

                            # Không ghi đè ô bằng dữ liệu rỗng
                            if value.strip():
                                sgcl_write_cell(
                                    ws, start_row, start_col, value
                                )
                                written_count += 1

                        # ---------------------------------
                        # DỮ LIỆU BẢNG
                        # ---------------------------------
                        elif mapping["type"] == "table":
                            table = mapping.get("table") or {}

                            headers = table.get("headers", [])
                            rows = table.get("rows", [])
                            column_map = mapping.get("column_map", {})

                            for index, header in enumerate(headers):
                                column_letter = str(
                                    column_map.get(
                                        header,
                                        openpyxl.utils.get_column_letter(
                                            start_col + index
                                        )
                                    )
                                ).strip().upper()

                                if not re.fullmatch(
                                    r"[A-Z]{1,3}",
                                    column_letter
                                ):
                                    raise ValueError(
                                        f"Cột Excel không hợp lệ: {column_letter}"
                                    )

                                column_number = (
                                    openpyxl.utils.column_index_from_string(
                                        column_letter
                                    )
                                )

                                for row_index, row_data in enumerate(rows):
                                    value = (
                                        row_data[index]
                                        if index < len(row_data)
                                        else ""
                                    )

                                    if str(value).strip():
                                        sgcl_write_cell(
                                            ws,
                                            start_row + row_index,
                                            column_number,
                                            value
                                        )
                                        written_count += 1

                    # -------------------------------------
                    # LƯU THÀNH FILE RIÊNG
                    # -------------------------------------
                    output = BytesIO()
                    wb.save(output)
                    wb.close()

                    ship["excel_output_bytes"] = output.getvalue()

                    safe_name = re.sub(
                        r"[^A-Za-z0-9_-]+",
                        "_",
                        ship["name"]
                    ).strip("_")

                    safe_voyage = re.sub(
                        r"[^A-Za-z0-9_-]+",
                        "_",
                        ship["voyage_no"]
                    ).strip("_")

                    ship["excel_output_name"] = (
                        f"SGCL_NSW_{safe_name}_"
                        f"{ship['imo']}_{safe_voyage}.xlsx"
                    )

                    ship["step_status"]["5"] = "done"
                    ship["status"] = "Hoàn thành"

                    st.success(
                        f"Đã tạo file Excel NSW thành công! "
                        f"Đã ghi {written_count} ô."
                    )

                    if skipped_fields:
                        st.info(
                            "Không xuất (chưa chọn ô): "
                            + ", ".join(skipped_fields)
                        )

                    save_all_ships_to_db()

                except Exception as e:
                    st.error(
                        f"Lỗi khi tạo file Excel: {e}"
                    )

            # -----------------------------------------
            # TẢI FILE ĐÃ XUẤT
            # -----------------------------------------
            if ship.get("excel_output_bytes"):
                st.divider()
                st.markdown("#### 📁 File Excel đã tạo")

                st.download_button(
                    label="⬇️ Tải xuống file Excel NSW",
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
                    key=f"s5_download_{ship['id']}"
                )

                st.caption(
                    "File kết quả được tạo từ bản sao của biểu mẫu. "
                    "TemplateNSW.xlsx gốc không bị thay đổi."
                )

        st.divider()

        if st.button(
            "← Quay lại Bước 4",
            use_container_width=True,
            key=f"s5_back_{ship['id']}"
        ):
            ship["current_step"] = 4
            save_all_ships_to_db()
            st.rerun()

# Đồng bộ dữ liệu hồ sơ sau mỗi lần render ứng dụng.
save_all_ships_to_db()
