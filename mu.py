import io
import json
import os
import re
import tempfile
import time
import zipfile
import requests
import streamlit as st
from bs4 import BeautifulSoup

# Import tkinter an toàn cho cả môi trường Local và Cloud (Linux không có GUI)
try:
    import tkinter as tk
    from tkinter import filedialog
except ImportError:
    tk = None

# --- Kiểm tra và nạp thư viện bổ trợ ---
try:
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Inches
except ImportError:
    os.system("pip install python-docx")
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Inches

try:
    import pypandoc
except ImportError:
    os.system("pip install pypandoc")
    import pypandoc

# ==========================================
# CẤU HÌNH THƯ MỤC LƯU & API KEY
# ==========================================
DEFAULT_OUTPUT_DIR = os.path.join(os.getcwd(), "downloads_mineru")
DEFAULT_API_KEY = "sk-IDb81Oj2W6pHrODooHN0xtKTxEXNzipsnZP6OxAqAl65Kz9O"

if "output_dir" not in st.session_state:
    st.session_state.output_dir = DEFAULT_OUTPUT_DIR

if "api_key" not in st.session_state:
    st.session_state.api_key = DEFAULT_API_KEY

if "edit_key_mode" not in st.session_state:
    st.session_state.edit_key_mode = False

os.makedirs(st.session_state.output_dir, exist_ok=True)

# ==========================================
# HÀM MỞ CỬA SỔ CHỌN FOLDER NATIVE (Chỉ chạy khi có GUI local)
# ==========================================
def select_folder():
    if tk:
        try:
            root = tk.Tk()
            root.withdraw()
            root.wm_attributes("-topmost", 1)
            folder_selected = filedialog.askdirectory(master=root)
            root.destroy()
            return folder_selected
        except Exception:
            return None
    return None

# ==========================================
# HÀM UPLOAD FILE TẠM ĐỂ GỬI API MINERU
# ==========================================
def upload_temp_file(uploaded_file):
    upload_url = "https://catbox.moe/user/api.php"
    files = {
        "fileToUpload": (
            uploaded_file.name,
            uploaded_file.getvalue(),
            uploaded_file.type,
        )
    }
    try:
        res = requests.post(
            upload_url, data={"reqtype": "fileupload"}, files=files, timeout=60
        )
        if res.status_code == 200 and res.text.startswith("http"):
            return res.text.strip()
    except Exception as e:
        st.error(f"Lỗi upload file gửi API: {e}")
    return None

# ==========================================
# CÁC HÀM XỬ LÝ DỮ LIỆU & PARSER LAYOUT.JSON
# ==========================================
def format_latex_string(latex_str):
    if not latex_str:
        return ""
    latex_clean = str(latex_str).strip()
    if latex_clean.startswith("$") and latex_clean.endswith("$"):
        latex_clean = latex_clean[1:-1].strip()
    latex_clean = latex_clean.replace("\n", " ")
    return f"${latex_clean}$"

def get_image_bytes(img_filename, images_dict, local_img_dir=None):
    if not img_filename:
        return None
    clean_name = os.path.basename(img_filename)
    
    # Ưu tiên 1: Lấy từ images_dict (nếu đọc file ZIP)
    if clean_name in images_dict:
        return io.BytesIO(images_dict[clean_name])
    
    # Ưu tiên 2: Tự động đọc trực tiếp từ thư mục trên máy tính (Không cần upload ảnh!)
    if local_img_dir and os.path.exists(local_img_dir):
        local_path = os.path.join(local_img_dir, clean_name)
        if os.path.exists(local_path):
            with open(local_path, "rb") as f:
                return io.BytesIO(f.read())
    return None

def process_html_table_md(table_html, images_dict, temp_dir=None, local_img_dir=None):
    soup = BeautifulSoup(table_html, "html.parser")
    rows = soup.find_all("tr")
    if not rows:
        return ""

    md_table = []
    headers_parsed = False
    img_counter = 0

    for tr in rows:
        cells = tr.find_all(["td", "th"])
        row_content = []
        for cell in cells:
            cell_text = ""
            for node in cell.children:
                if node.name == "eq":
                    cell_text += f" {format_latex_string(node.get_text())} "
                elif node.name == "img":
                    img_src = node.get("src")
                    if img_src and temp_dir:
                        img_stream = get_image_bytes(img_src, images_dict, local_img_dir)
                        if img_stream:
                            img_counter += 1
                            img_path = os.path.join(temp_dir, f"tbl_img_{img_counter}.png")
                            with open(img_path, "wb") as f:
                                f.write(img_stream.getvalue())
                            cell_text += f" ![]({img_path}) "
                elif node.name is None:
                    cell_text += str(node).strip()
            row_content.append(cell_text.strip().replace("\n", " "))

        md_row = "| " + " | ".join(row_content) + " |"
        md_table.append(md_row)

        if not headers_parsed:
            separator = "| " + " | ".join(["---"] * len(row_content)) + " |"
            md_table.append(separator)
            headers_parsed = True

    return "\n".join(md_table) + "\n\n"

def convert_json_to_markdown(json_data, images_dict, temp_dir=None, local_img_dir=None):
    md_lines = []
    img_counter = 0

    if not json_data:
        return ""

    # TH1: Parse dạng Mảng phẳng
    if isinstance(json_data, list):
        is_flat_list = True
        for item in json_data[:3]:
            if isinstance(item, dict) and ("pdf_info" in item or "page_info" in item):
                is_flat_list = False
                break

        if is_flat_list:
            for item in json_data:
                if not isinstance(item, dict):
                    continue
                b_type = item.get("type", "")
                text_content = item.get("text", item.get("content", item.get("body", "")))

                if b_type in ["text", "title", "header", "footer", "paragraph"]:
                    if text_content:
                        md_lines.append(f"{text_content.strip()}\n\n")

                elif b_type in ["inline_equation", "interline_equation", "equation", "math"]:
                    if text_content:
                        md_lines.append(f" {format_latex_string(text_content)} \n\n")

                elif b_type in ["image", "chart"]:
                    img_path = item.get("img_path", item.get("image_path", item.get("path", "")))
                    if img_path:
                        img_stream = get_image_bytes(img_path, images_dict, local_img_dir)
                        if img_stream and temp_dir:
                            img_counter += 1
                            local_img_path = os.path.join(temp_dir, f"img_{img_counter}.png")
                            with open(local_img_path, "wb") as f:
                                f.write(img_stream.getvalue())
                            md_lines.append(f"![Hình ảnh]({local_img_path})\n\n")

                elif b_type == "table":
                    table_html = item.get("table_html", item.get("html", ""))
                    if table_html:
                        md_lines.append(process_html_table_md(table_html, images_dict, temp_dir, local_img_dir))
                    elif text_content:
                        md_lines.append(f"{text_content.strip()}\n\n")

            if md_lines:
                return "".join(md_lines)

    # TH2: Parse cấu trúc layout.json
    pages = []
    if isinstance(json_data, list):
        for item in json_data:
            if isinstance(item, dict):
                if "pdf_info" in item:
                    pages.extend(item["pdf_info"])
                else:
                    pages.append(item)
    elif isinstance(json_data, dict):
        if "pdf_info" in json_data:
            pages = json_data["pdf_info"]
        else:
            pages = [json_data]

    for page in pages:
        if not isinstance(page, dict):
            continue

        para_blocks = page.get("para_blocks", page.get("blocks", page.get("preproc_blocks", [])))
        if not isinstance(para_blocks, list):
            continue

        for block in para_blocks:
            if not isinstance(block, dict):
                continue
            b_type = block.get("type", "text")

            if b_type in ["text", "title", "header", "footer", "paragraph"]:
                p_text = ""
                lines = block.get("lines", [])
                for line in lines:
                    if not isinstance(line, dict):
                        continue
                    for span in line.get("spans", []):
                        if not isinstance(span, dict):
                            continue
                        span_type = span.get("type")
                        content = span.get("content", span.get("text", ""))

                        if span_type in ["inline_equation", "interline_equation", "equation", "math"]:
                            p_text += f" {format_latex_string(content)} "
                        else:
                            if re.match(r"^Bài\s+\d+", content.strip()):
                                p_text += f"**{content}**"
                            else:
                                p_text += content
                if p_text.strip():
                    md_lines.append(p_text.strip() + "\n\n")

            elif b_type in ["image", "chart"]:
                blocks_to_check = block.get("blocks", [block])
                for sub_b in blocks_to_check:
                    if not isinstance(sub_b, dict):
                        continue
                    for line in sub_b.get("lines", []):
                        if not isinstance(line, dict):
                            continue
                        for span in line.get("spans", []):
                            if not isinstance(span, dict):
                                continue
                            img_path = span.get("image_path", span.get("img_path", ""))
                            if img_path:
                                img_stream = get_image_bytes(img_path, images_dict, local_img_dir)
                                if img_stream and temp_dir:
                                    img_counter += 1
                                    local_img_path = os.path.join(temp_dir, f"img_{img_counter}.png")
                                    with open(local_img_path, "wb") as f:
                                        f.write(img_stream.getvalue())
                                    md_lines.append(f"![Hình ảnh]({local_img_path})\n\n")

            elif b_type == "table":
                blocks_to_check = block.get("blocks", [block])
                for sub_b in blocks_to_check:
                    if not isinstance(sub_b, dict):
                        continue
                    for line in sub_b.get("lines", []):
                        if not isinstance(line, dict):
                            continue
                        for span in line.get("spans", []):
                            if not isinstance(span, dict):
                                continue
                            table_html = span.get("html", span.get("table_html", ""))
                            if table_html:
                                md_lines.append(process_html_table_md(table_html, images_dict, temp_dir, local_img_dir))

    return "".join(md_lines)

def convert_json_to_docx_pandoc_bytes(json_data, images_dict, local_img_dir=None):
    with tempfile.TemporaryDirectory() as temp_dir:
        md_text = convert_json_to_markdown(json_data, images_dict, temp_dir=temp_dir, local_img_dir=local_img_dir)

        if not md_text.strip():
            raise ValueError("Nội dung Markdown sau khi giải mã bị rỗng. Vui lòng kiểm tra lại file JSON!")

        output_docx_path = os.path.join(temp_dir, "output_native.docx")
        current_cwd = os.getcwd()
        try:
            os.chdir(temp_dir)
            pypandoc.convert_text(
                source=md_text,
                format="markdown",
                to="docx",
                outputfile=output_docx_path,
                extra_args=["--mathjax"],
            )
            with open(output_docx_path, "rb") as f:
                return f.read()
        finally:
            os.chdir(current_cwd)

def extract_zip_and_get_data(zip_bytes):
    images_dict = {}
    json_files = {}

    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as z:
        for filename in z.namelist():
            if "images/" in filename and not filename.endswith("/"):
                img_name = os.path.basename(filename)
                images_dict[img_name] = z.read(filename)
            elif filename.endswith(".json") and not filename.startswith("__MACOSX"):
                try:
                    json_files[filename] = json.loads(z.read(filename).decode("utf-8"))
                except Exception:
                    pass

    selected_file = None
    for fname in json_files:
        if "layout.json" in fname.lower():
            selected_file = fname
            break

    if not selected_file:
        for fname in json_files:
            if "middle" in fname or "model" in fname:
                selected_file = fname
                break

    if not selected_file:
        for fname in json_files:
            if "content_list" in fname:
                selected_file = fname
                break

    if not selected_file and json_files:
        selected_file = list(json_files.keys())[0]

    json_data = json_files.get(selected_file, {}) if selected_file else {}
    return json_data, images_dict

# ==========================================
# GIAO DIỆN STREAMLIT
# ==========================================
st.set_page_config(page_title="MinerU Local Workflow Manager", page_icon="💾", layout="wide")
st.title("💾 MinerU Extraction & Local Auto-Save Converter")

# SIDEBAR
with st.sidebar:
    st.header("⚙️ Cấu hình Hệ Thống")
    
    # 1. API KEY
    st.subheader("🔑 API Key")
    col_input, col_btn = st.columns([3, 1.2])
    with col_input:
        new_key_input = st.text_input(
            "API Key",
            value=st.session_state.api_key,
            type="password",
            disabled=not st.session_state.edit_key_mode,
            label_visibility="collapsed",
        )
    with col_btn:
        if not st.session_state.edit_key_mode:
            if st.button("✏️ Đổi", use_container_width=True):
                st.session_state.edit_key_mode = True
                st.rerun()
        else:
            if st.button("💾 Lưu", use_container_width=True, type="primary"):
                st.session_state.api_key = new_key_input.strip()
                st.session_state.edit_key_mode = False
                st.rerun()

    st.markdown("---")
    
    # 2. CHỌN THƯ MỤC MẶC ĐỊNH LƯU ZIP
    st.subheader("📂 Thư mục lưu ZIP")
    st.text_input("Đường dẫn hiện tại:", value=st.session_state.output_dir, key="dir_display", disabled=True)
    
    col_dir1, col_dir2 = st.columns([1, 1])
    with col_dir1:
        if st.button("📁 Chọn Thư Mục", use_container_width=True):
            chosen_dir = select_folder()
            if chosen_dir:
                st.session_state.output_dir = chosen_dir
                os.makedirs(chosen_dir, exist_ok=True)
                st.rerun()
    with col_dir2:
        if st.button("🔄 Mặc định", use_container_width=True):
            st.session_state.output_dir = DEFAULT_OUTPUT_DIR
            st.rerun()

tab1, tab2 = st.tabs([
    "🚀 Trích xuất API (Tự Lưu ZIP về Máy & Tạo Word)",
    "📦 Convert File Sẵn Có (ZIP hoặc File JSON + Folder Ảnh)",
])

# --- TAB 1: TRÍCH XUẤT API VÀ LƯU THẲNG VỀ MÁY TÍNH ---
with tab1:
    uploaded_file = st.file_uploader(
        "Tải lên file tài liệu (PDF, PNG, JPG, DOCX...):",
        type=["pdf", "docx", "doc", "pptx", "ppt", "xlsx", "xls", "png", "jpg", "jpeg"],
        key="api_file",
    )

    if uploaded_file and st.button("🚀 Bắt đầu trích xuất", type="primary", use_container_width=True):
        token = st.session_state.api_key.strip()
        headers = {"Content-Type": "application/json", "Authorization": f"Bearer {token}"}
        base_name = uploaded_file.name.rsplit(".", 1)[0]

        with st.spinner("Đang tải file gửi lên MinerU..."):
            file_url = upload_temp_file(uploaded_file)
            if not file_url:
                st.stop()

        with st.spinner("Đang khởi tạo Task MinerU..."):
            payload = {"url": file_url, "model_version": "vlm", "is_ocr": True}
            res = requests.post("https://mineru.net/api/v4/extract/task", headers=headers, json=payload)
            if res.status_code != 200 or res.json().get("code") != 0:
                st.error(f"Lỗi khởi tạo: {res.text}")
                st.stop()
            task_id = res.json()["data"]["task_id"]

        query_url = f"https://mineru.net/api/v4/extract/task/{task_id}"
        progress_bar = st.progress(0)
        task_done = False
        final_data = {}

        while not task_done:
            res_status = requests.get(query_url, headers=headers)
            if res_status.status_code == 200:
                data = res_status.json().get("data", {})
                state = data.get("state")
                if state == "done":
                    progress_bar.progress(100)
                    final_data = data
                    task_done = True
                elif state == "failed":
                    st.error("Xử lý thất bại!")
                    st.stop()
                else:
                    progress_bar.progress(50)
                    time.sleep(3)

        st.info("⚡ Đang lưu file ZIP về máy tính và đóng gói Word...")
        zip_url = final_data.get("full_zip_url")
        images_dict = {}
        json_data = {}
        saved_zip_path = ""

        if zip_url:
            zip_res = requests.get(zip_url)
            if zip_res.status_code == 200:
                zip_bytes = zip_res.content
                saved_zip_path = os.path.join(st.session_state.output_dir, f"{base_name}_mineru.zip")
                try:
                    with open(saved_zip_path, "wb") as f:
                        f.write(zip_bytes)
                except Exception:
                    pass

                json_data, images_dict = extract_zip_and_get_data(zip_bytes)

        if json_data:
            st.success(f"🎉 Hoàn tất! File ZIP đã được tự động xử lý!")
            if saved_zip_path:
                st.code(f"Location ZIP: {saved_zip_path}", language="bash")

            st.markdown("---")
            col1, col2 = st.columns(2)

            with col1:
                try:
                    with st.spinner("Đang dựng Word Native Math..."):
                        docx_pandoc = convert_json_to_docx_pandoc_bytes(json_data, images_dict)
                    st.download_button(
                        label="📐 Tải Word (Native Math Equation + Ảnh)",
                        data=docx_pandoc,
                        file_name=f"{base_name}_NativeMath.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        type="primary",
                        use_container_width=True,
                    )
                except Exception as e:
                    st.error(f"Lỗi tạo Word: {e}")

            with col2:
                md_content = convert_json_to_markdown(json_data, images_dict)
                st.download_button(
                    label="📄 Tải File Markdown (.md)",
                    data=md_content,
                    file_name=f"{base_name}.md",
                    mime="text/markdown",
                    use_container_width=True,
                )

# --- TAB 2: CONVERT TỪ MÁY TÍNH (ZIP HOẶC JSON ĐƠN) ---
with tab2:
    st.subheader("Chọn Phương Thức Nạp Dữ Liệu")
    input_mode = st.radio("Chế độ:", ["📦 Chọn File ZIP", "📄 Chọn File JSON (Tự đọc ảnh từ Folder)"], horizontal=True)

    json_data = {}
    images_dict = {}
    local_img_dir = None
    base_name = "Converted_Doc"

    if input_mode == "📦 Chọn File ZIP":
        zip_file = st.file_uploader("Kéo thả File .zip từ MinerU:", type=["zip"], key="offline_zip")
        if zip_file:
            base_name = zip_file.name.rsplit(".", 1)[0]
            json_data, images_dict = extract_zip_and_get_data(zip_file.getvalue())

    else:
        json_file = st.file_uploader("1. Tải lên file JSON (ví dụ: layout.json):", type=["json"], key="offline_json")
        
        st.markdown("**2. Đường dẫn thư mục chứa Ảnh (Local hoặc trên server):**")
        col_path, col_btn_path = st.columns([3, 1])
        
        if "local_img_path" not in st.session_state:
            st.session_state.local_img_path = ""

        with col_path:
            local_img_dir = st.text_input("Đường dẫn thư mục images:", value=st.session_state.local_img_path, placeholder="E:\\AppChuyenDoiTex\\downloads_mineru\\images")
            st.session_state.local_img_path = local_img_dir

        with col_btn_path:
            st.write("") 
            st.write("")
            if st.button("📁 Chọn Folder Ảnh"):
                chosen_img_dir = select_folder()
                if chosen_img_dir:
                    st.session_state.local_img_path = chosen_img_dir
                    st.rerun()

        if json_file:
            base_name = json_file.name.rsplit(".", 1)[0]
            try:
                json_data = json.loads(json_file.getvalue().decode("utf-8"))
            except Exception as e:
                st.error(f"Lỗi đọc JSON: {e}")

    if json_data:
        st.success("✅ Đã nạp thành công dữ liệu JSON!")
        st.markdown("---")

        col1, col2 = st.columns(2)

        with col1:
            try:
                with st.spinner("Đang tạo Word Native Equation..."):
                    docx_pandoc = convert_json_to_docx_pandoc_bytes(json_data, images_dict, local_img_dir=local_img_dir)
                st.download_button(
                    label="📐 Tải Word (Native Math Equation + Ảnh)",
                    data=docx_pandoc,
                    file_name=f"{base_name}_NativeMath.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary",
                    use_container_width=True,
                )
            except Exception as e:
                st.error(f"Lỗi tạo Word: {e}")

        with col2:
            md_content = convert_json_to_markdown(json_data, images_dict, local_img_dir=local_img_dir)
            st.download_button(
                label="📄 Tải File Markdown (.md)",
                data=md_content,
                file_name=f"{base_name}.md",
                mime="text/markdown",
                use_container_width=True,
            )