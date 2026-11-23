
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
    """Đồng bộ toàn bộ hồ sơ trong session xuống SQLite."""
    for item in st.session_state.get("ships", []):
        save_ship_to_db(item)


# Khởi tạo cơ sở dữ liệu khi ứng dụng chạy
init_database()
# =========================================================
# 1. CẤU HÌNH LƯU TRỮ
# =========================================================

BASE_UPLOAD_DIR = Path("uploads")
BASE_UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
ALLOWED_EXTENSIONS = {".pdf"}
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

        else:
            st.warning("Chỉ hỗ trợ xem trước file PDF.")

    except Exception as e:
        st.error(f"Không thể xem trước chứng từ: {str(e)}")


# =========================================================
# 3. CÁC HÀM HỖ TRỢ PDF/OCR - BƯỚC 3
# =========================================================

def sgcl_get_document_bytes(doc):
    """Đọc chứng từ đã tải lên ở Bước 2."""
    file_path = Path(doc["path"])

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
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
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
        lang="eng",
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
    if CONFIG_PATH.exists():
        try:
            return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        except Exception:
            pass

    CONFIG_PATH.write_text(
        json.dumps(DEFAULT_FIELD_CONFIG, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )
    return json.loads(json.dumps(DEFAULT_FIELD_CONFIG, ensure_ascii=False))


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


def sgcl_get_pdf_pages(file_bytes):
    """Đọc text layer của tất cả các trang PDF."""
    pages = []

    with fitz.open(stream=file_bytes, filetype="pdf") as pdf:
        for index, page in enumerate(pdf, start=1):
            pages.append({
                "page": index,
                "text": page.get_text("text") or ""
            })

    return pages


def sgcl_get_all_pdf_text(file_bytes):
    """
    Ưu tiên đọc text layer.
    Chỉ OCR những trang gần như không có văn bản.
    Nếu OCR thất bại thì báo lỗi rõ ràng.
    """

    pages = sgcl_get_pdf_pages(file_bytes)

    for item in pages:
        if len(item["text"].strip()) < 15:
            page_number = item["page"]

            try:
                import pytesseract

                with fitz.open(
                    stream=file_bytes,
                    filetype="pdf"
                ) as pdf:

                    page = pdf[page_number - 1]

                    pix = page.get_pixmap(
                        matrix=fitz.Matrix(2, 2),
                        alpha=False
                    )

                    image = Image.open(
                        BytesIO(pix.tobytes("png"))
                    ).convert("RGB")

                item["text"] = pytesseract.image_to_string(
                    image,
                    lang="eng",
                    config="--oem 3 --psm 6"
                )

                item["source_type"] = "OCR"

            except Exception as e:
                raise RuntimeError(
                    f"Không thể OCR trang {page_number}. "
                    f"Vui lòng kiểm tra pytesseract và Tesseract OCR. "
                    f"Chi tiết lỗi: {str(e)}"
                ) from e

        else:
            item["source_type"] = "Text layer"

    return pages


def sgcl_extract_single(pages, field_name):
    """Chỉ lấy giá trị khi tìm được nhãn tương ứng."""
    aliases = {
        "Arrival draft (F/A)": ["Arrival draft", "Draft F/A"],
        "Air draft": ["Air draft"],
        "Last port of call": ["Last port of call", "Last port"],
        "Port of arrival": ["Port of arrival", "Arrival port"],
        "Departure from last port": ["Departure from last port"],
        "Time of arrival": ["Time of arrival", "Arrival time"],
        "All fast": ["All fast", "Allfast"],
        "P.O.B": ["P.O.B", "Pilot on board"],
        "F.O": ["Fuel oil", "F.O"],
        "M.D.O": ["Diesel oil", "M.D.O"],
        "FreshWater": ["Fresh water", "FreshWater"],
        "Berthing Position": ["Berthing position", "Berth"],
        "Port of discharge": ["Port of discharge"],
        "B/L No.": ["B/L No", "Bill of lading No"],
        "Gross weight": ["Gross weight"],
        "Port of Loading": ["Port of loading"],
        "Port of Discharge": ["Port of discharge"],
        "Arriving from": ["Arriving from"],
        "Sailing to": ["Sailing to"],
        "Issued at": ["Issued at"],
    }

    search_labels = [field_name] + aliases.get(field_name, [])
    search_keys = [sgcl_normalize(x) for x in search_labels]

    for page in pages:
        lines = [
            re.sub(r"\s+", " ", line).strip()
            for line in page["text"].splitlines()
        ]

        for index, line in enumerate(lines):
            normalized_line = sgcl_normalize(line)

            for label, key in zip(search_labels, search_keys):
                if not key or key not in normalized_line:
                    continue

                value = ""

                if ":" in line:
                    value = line.split(":", 1)[1].strip()
                elif "：" in line:
                    value = line.split("：", 1)[1].strip()
                else:
                    match = re.search(re.escape(label), line, re.I)
                    if match:
                        value = line[match.end():].strip(" :-–—")

                if not value and index + 1 < len(lines):
                    next_line = lines[index + 1].strip()
                    if next_line and len(next_line) < 250:
                        value = next_line

                if value:
                    return {
                        "value": value,
                        "page": page["page"],
                        "found": True
                    }

    return {"value": "", "page": None, "found": False}

def sgcl_extract_table(pages, headers):
    """
    Tách bảng cơ bản từ PDF có text layer.
    Không tự suy đoán hoặc tạo dữ liệu khi không tìm thấy.
    """

    rows = []
    source_page = None

    if not headers:
        return {
            "headers": [],
            "rows": [],
            "page": None,
            "found": False
        }

    normalized_headers = [
        sgcl_normalize(header)
        for header in headers
    ]

    for page in pages:
        lines = page["text"].splitlines()

        for line in lines:
            line = line.strip()

            if not line:
                continue

            # Tách cột theo tab hoặc khoảng trắng lớn
            parts = [
                re.sub(r"\s+", " ", x).strip()
                for x in re.split(r"\t+|\s{2,}", line)
                if x.strip()
            ]

            # Bỏ qua dòng không đủ dữ liệu cột
            if len(parts) < 2:
                continue

            # Bỏ qua dòng tiêu đề
            normalized_parts = [
                sgcl_normalize(part)
                for part in parts
            ]

            is_header = all(
                header in normalized_parts
                for header in normalized_headers
            )

            if is_header:
                continue

            # Chỉ nhận dòng có số cột gần với cấu trúc bảng
            if len(parts) > len(headers) + 2:
                continue

            # Chuẩn hóa số cột, không tự tạo giá trị
            row = parts[:len(headers)]

            while len(row) < len(headers):
                row.append("")

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
                save_all_ships_to_db()
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
                    save_ship_to_db(new_ship)
                    st.session_state.show_form = False

                    st.success("Đã tạo hồ sơ thành công!")
                    time.sleep(0.5)
                    save_all_ships_to_db()
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

                    save_all_ships_to_db()
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

        uploaded_files = st.file_uploader(
            "Chọn một hoặc nhiều chứng từ",
            type=["pdf"],
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
                        save_all_ships_to_db()
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
                    key=f"s3_config_editor_{ship['id']}_{sgcl_normalize(selected_config_type)}"
                )

                aliases_text = st.text_input(
                    "Từ khóa nhận diện tên file, cách nhau bằng dấu phẩy",
                    value=", ".join(
                        config[selected_config_type].get("aliases", [])
                    ),
                    key=f"s3_aliases_{ship['id']}"
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

                    st.success("Đã lưu cấu hình trường dữ liệu.")
                    save_all_ships_to_db()
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

            file_path = Path(selected_doc["path"])

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
                        pix = page.get_pixmap(
                            matrix=fitz.Matrix(1.5, 1.5),
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
                                            field["name"]
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
                                            headers
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

                                    st.success(
                                        "Đã trích xuất xong. "
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

                            normalized_rows = [
                                list(row) + [""] * max(0, len(headers) - len(row))
                                for row in rows
                            ]

                            df = pd.DataFrame(
                                normalized_rows,
                                columns=headers
                            )

                            edited_df = st.data_editor(
                                df,
                                num_rows="dynamic",
                                use_container_width=True,
                                hide_index=True,
                                key=f"s3_table_{ship['id']}_{rec['id']}"
                            )

                            new_rows = (
                                edited_df.fillna("")
                                .astype(str)
                                .values.tolist()
                            )

                            if new_rows != rows:
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
        extraction_ids = {
            rec["id"] for rec in ship["field_extractions"]
        }

        existing_ids = {
            item.get("source_record_id")
            for item in ship["mappings"]
        }

        for rec in ship["field_extractions"]:
            if rec["id"] not in existing_ids:
                mapping = {
                    "id": str(uuid.uuid4()),
                    "source_record_id": rec["id"],
                    "field": rec["field"],
                    "type": rec["type"],
                    "document": rec.get("document", ""),
                    "source_type": rec.get("source_type", "extracted"),
                    "value": (
                        rec.get("result", {}).get("value", "")
                        if rec["type"] == "single" else ""
                    ),
                    "table": (
                        rec.get("result", {})
                        if rec["type"] == "table" else None
                    ),
                    "sheet": NSW_SHEETS[0],
                    "cell": "A1",
                    "column_map": {}
                }

                ship["mappings"].append(mapping)

        # Đồng bộ các chỉnh sửa từ Bước 3
        extraction_lookup = {
            rec["id"]: rec
            for rec in ship["field_extractions"]
        }

        for mapping in ship["mappings"]:
            source_id = mapping.get("source_record_id")

            if source_id and source_id in extraction_lookup:
                rec = extraction_lookup[source_id]

                mapping["field"] = rec["field"]
                mapping["type"] = rec["type"]
                mapping["document"] = rec.get("document", "")
                mapping["source_type"] = rec.get(
                    "source_type", "extracted"
                )

                if rec["type"] == "single":
                    mapping["value"] = rec["result"].get("value", "")
                else:
                    mapping["table"] = rec.get("result", {})

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

        if not ship["mappings"]:
            st.info(
                "Chưa có dữ liệu. Bà có thể nhập thủ công ở phần bên dưới."
            )

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
                        st.write(
                            mapping.get("value", "") or "Chưa có giá trị"
                        )

                    else:
                        table = mapping.get("table") or {}
                        headers = table.get("headers", [])
                        rows = table.get("rows", [])

                        if headers:
                            df = pd.DataFrame(
                                rows,
                                columns=headers
                            )
                            st.dataframe(
                                df,
                                use_container_width=True,
                                hide_index=True
                            )

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
                        "Ô đích / ô bắt đầu",
                        value=mapping.get("cell", "A1"),
                        key=f"s4_cell_{ship['id']}_{mapping['id']}"
                    ).strip().upper()

                    cell_valid = True

                    try:
                        openpyxl.utils.cell.coordinate_from_string(
                            selected_cell
                        )
                    except Exception:
                        cell_valid = False
                        st.error("Địa chỉ ô không hợp lệ. Ví dụ: B5, C10.")

                    mapping["sheet"] = selected_sheet
                    mapping["cell"] = selected_cell

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
                        headers = table.get("headers", [])

                        st.caption(
                            "Chọn cột Excel tương ứng cho từng trường bảng."
                        )

                        try:
                            start_col_letter, start_row = (
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

                        column_map = mapping.get("column_map", {})

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

                            column_map[header] = column_value

                        mapping["column_map"] = column_map

                        # Cho phép chỉnh sửa bảng trước khi xác nhận
                        rows = table.get("rows", [])
                        normalized_rows = [
                            list(row) + [""] * max(0, len(headers) - len(row))
                            for row in rows
                        ]

                        edit_df = pd.DataFrame(
                            normalized_rows,
                            columns=headers
                        )

                        edited_df = st.data_editor(
                            edit_df,
                            num_rows="dynamic",
                            use_container_width=True,
                            hide_index=True,
                            key=f"s4_table_edit_{ship['id']}_{mapping['id']}"
                        )

                        new_rows = (
                            edited_df.fillna("")
                            .astype(str)
                            .values.tolist()
                        )

                        if new_rows != rows:
                            mapping["table"]["rows"] = new_rows
                            if mapping["source_type"] == "extracted":
                                mapping["source_type"] = "manually_edited"
                            sgcl_invalidate_review(ship)

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

        with st.form(f"s4_manual_form_{ship['id']}"):
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
                height=80
            )

            manual_sheet = st.selectbox(
                "Sheet Excel",
                NSW_SHEETS,
                key=f"s4_manual_sheet_{ship['id']}"
            )

            manual_cell = st.text_input(
                "Ô đích / ô bắt đầu",
                value="A1"
            )

            add_manual = st.form_submit_button(
                "➕ Thêm dữ liệu",
                type="primary"
            )

            if add_manual:
                if not manual_name.strip():
                    st.error("Vui lòng nhập tên trường.")
                else:
                    try:
                        openpyxl.utils.cell.coordinate_from_string(
                            manual_cell.strip().upper()
                        )

                        is_table = manual_type == "Bảng nhiều dòng"

                        new_mapping = {
                            "id": str(uuid.uuid4()),
                            "source_record_id": None,
                            "field": manual_name.strip(),
                            "type": "table" if is_table else "single",
                            "document": "",
                            "source_type": "manually_entered",
                            "value": "" if is_table else manual_value,
                            "table": (
                                {
                                    "headers": [manual_name.strip()],
                                    "rows": []
                                }
                                if is_table else None
                            ),
                            "sheet": manual_sheet,
                            "cell": manual_cell.strip().upper(),
                            "column_map": {}
                        }

                        ship["mappings"].append(new_mapping)
                        ship["manual_entries"].append(
                            new_mapping["id"]
                        )

                        sgcl_invalidate_review(ship)
                        st.success("Đã thêm dữ liệu thủ công.")
                        save_all_ships_to_db()
                        st.rerun()

                    except Exception:
                        st.error(
                            "Ô Excel không hợp lệ. Ví dụ: A1, B5, C10."
                        )

        # ---------------------------------------------
        # KIỂM TRA TRƯỚC KHI XÁC NHẬN
        # ---------------------------------------------
        st.divider()
        st.markdown("#### Kiểm tra trước khi xác nhận")

        invalid_mappings = [
            m for m in ship["mappings"]
            if m.get("_invalid")
        ]

        missing_values = [
            m for m in ship["mappings"]
            if m["type"] == "single"
            and not str(m.get("value", "")).strip()
        ]

        if invalid_mappings:
            st.error(
                f"Có {len(invalid_mappings)} địa chỉ ô Excel không hợp lệ."
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
                not ship["mappings"]
                or bool(invalid_mappings)
            )
        ):
            ship["review_confirmed"] = True
            ship["step_status"]["4"] = "done"
            ship["status"] = "Đã xác nhận"
            ship["current_step"] = 5
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
                    "Ô bắt đầu": mapping.get("cell", ""),
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
                                ws[start_cell] = value

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
                                        ws.cell(
                                            row=start_row + row_index,
                                            column=column_number,
                                            value=value
                                        )

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
                        "Đã tạo file Excel NSW thành công!"
                    )

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
