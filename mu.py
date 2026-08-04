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

# Import các thư viện phụ thuộc
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches
import pypandoc

# Kiểm tra tkinter
try:
    import tkinter as tk
    from tkinter import filedialog
except ImportError:
    tk = None

# ==========================================
# CẤU HÌNH THƯ MỤC LƯU & API KEY
# ==========================================
DEFAULT_OUTPUT_DIR = os.path.join(tempfile.gettempdir(), "downloads_mineru")
DEFAULT_API_KEY = "sk-IDb81Oj2W6pHrODooHN0xtKTxEXNzipsnZP6OxAqAl65Kz9O"

if "output_dir" not in st.session_state:
    st.session_state.output_dir = DEFAULT_OUTPUT_DIR

if "api_key" not in st.session_state:
    st.session_state.api_key = DEFAULT_API_KEY

if "edit_key_mode" not in st.session_state:
    st.session_state.edit_key_mode = False

os.makedirs(st.session_state.output_dir, exist_ok=True)

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
    
    if clean_name in images_dict:
        return io.BytesIO(images_dict[clean_name])
    
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
            raise ValueError("Nội dung Markdown sau khi giải mã bị rỗng!")

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
st.set_page_config(page_title="MinerU Extraction & Converter", page_icon="💾", layout="wide")
st.title("🌐 MinerU Online Converter")

# SIDEBAR
with st.sidebar:
    st.header("⚙️ Cấu hình Hệ Thống")
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

tab1, tab2 = st.tabs([
    "🚀 Trích xuất API (MinerU Cloud)",
    "📦 Convert File Sẵn Có (File ZIP từ MinerU)",
])

# --- TAB 1: TRÍCH XUẤT API TRỰC TIẾP QUA MINERU UPLOAD ---
with tab1:
    uploaded_file = st.file_uploader(
        "Tải lên file tài liệu (PDF, PNG, JPG, DOCX...):",
        type=["pdf", "docx", "doc", "pptx", "ppt", "xlsx", "xls", "png", "jpg", "jpeg"],
        key="api_file",
    )

    btn_click = st.button("🚀 Bắt đầu trích xuất", type="primary", use_container_width=True)

    if btn_click:
        if not uploaded_file:
            st.warning("⚠️ Vui lòng chọn file trước khi bấm trích xuất!")
        else:
            token = st.session_state.api_key.strip()
            headers = {
                "Content-Type": "application/json",
                "Authorization": f"Bearer {token}"
            }
            base_name = uploaded_file.name.rsplit(".", 1)[0]

            status_box = st.empty()
            progress_bar = st.progress(0)

            # BƯỚC 1: Lấy URL upload từ MinerU API
            status_box.info("🔑 [1/4] Xin cấp quyền upload từ MinerU...")
            get_url_payload = {
                "files": [{"name": uploaded_file.name}]
            }
            
            try:
                res_url = requests.post(
                    "https://mineru.net/api/v4/file-urls",
                    headers=headers,
                    json=get_url_payload,
                    timeout=30
                )
                
                if res_url.status_code == 200 and res_url.json().get("code") == 0:
                    upload_info = res_url.json()["data"][0]
                    put_url = upload_info["upload_url"]
                    file_batch_id = upload_info.get("batch_id")
                    
                    progress_bar.progress(25)
                    # BƯỚC 2: Upload file trực tiếp lên Storage của MinerU
                    status_box.info("📤 [2/4] Đang upload file trực tiếp lên MinerU Cloud...")
                    put_res = requests.put(
                        put_url,
                        data=uploaded_file.getvalue(),
                        headers={"Content-Type": uploaded_file.type or "application/octet-stream"},
                        timeout=120
                    )
                    
                    if put_res.status_code in [200, 201]:
                        progress_bar.progress(50)
                        # BƯỚC 3: Tạo Task trích xuất
                        status_box.info("⚙️ [3/4] Đang kích hoạt nhiệm vụ trích xuất MinerU...")
                        extract_payload = {
                            "batch_id": file_batch_id,
                            "files": [{"name": uploaded_file.name, "is_ocr": True, "model_version": "vlm"}]
                        }
                        
                        task_res = requests.post(
                            "https://mineru.net/api/v4/extract/task",
                            headers=headers,
                            json=extract_payload,
                            timeout=30
                        )
                        
                        if task_res.status_code == 200 and task_res.json().get("code") == 0:
                            task_data = task_res.json().get("data", {})
                            
                            # Lấy task_id (xử lý cả trường hợp trả về mảng hoặc dict)
                            if isinstance(task_data, list) and len(task_data) > 0:
                                task_id = task_data[0].get("task_id")
                            elif isinstance(task_data, dict):
                                task_id = task_data.get("task_id")
                            else:
                                task_id = None

                            if not task_id:
                                status_box.error(f"❌ MinerU không trả về Task ID: {task_res.text}")
                                st.stop()

                            # BƯỚC 4: Vòng lặp chờ MinerU OCR
                            query_url = f"https://mineru.net/api/v4/extract/task/{task_id}"
                            task_done = False
                            final_data = {}
                            start_time = time.time()

                            while not task_done:
                                if time.time() - start_time > 180:
                                    status_box.error("⏱️ Quá thời gian chờ (Timeout 3 phút). Vui lòng thử lại!")
                                    break

                                status_box.info("🔄 [4/4] MinerU đang OCR và nhận diện công thức toán... Vui lòng chờ...")
                                try:
                                    res_status = requests.get(query_url, headers=headers, timeout=15)
                                    if res_status.status_code == 200:
                                        data = res_status.json().get("data", {})
                                        state = data.get("state")
                                        if state == "done":
                                            progress_bar.progress(80)
                                            final_data = data
                                            task_done = True
                                        elif state == "failed":
                                            status_box.error("❌ MinerU xử lý file thất bại!")
                                            break
                                except Exception:
                                    pass

                                time.sleep(3)

                            # BƯỚC 5: Tải ZIP & chuyển đổi thành Word
                            if task_done and final_data.get("full_zip_url"):
                                status_box.info("⚡ Tải gói dữ liệu và dựng file Word Native Math...")
                                zip_res = requests.get(final_data["full_zip_url"], timeout=60)
                                
                                if zip_res.status_code == 200:
                                    json_data, images_dict = extract_zip_and_get_data(zip_res.content)
                                    
                                    if json_data:
                                        progress_bar.progress(100)
                                        status_box.success("🎉 Trích xuất thành công!")
                                        st.markdown("---")
                                        
                                        col1, col2 = st.columns(2)
                                        with col1:
                                            try:
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
                                    else:
                                        status_box.error("❌ Không tìm thấy file JSON hợp lệ trong ZIP trả về.")
                                else:
                                    status_box.error("❌ Không thể tải ZIP từ MinerU.")
                        else:
                            status_box.error(f"❌ Không thể tạo Task MinerU: {task_res.text}")
                    else:
                        status_box.error(f"❌ Lỗi Upload trực tiếp lên MinerU (Status {put_res.status_code})")
                else:
                    status_box.error(f"❌ MinerU từ chối cấp Link Upload: {res_url.text}")
            except Exception as e:
                status_box.error(f"❌ Lỗi kết nối API: {e}")

# --- TAB 2: CONVERT TỪ ZIP ---
with tab2:
    st.subheader("Nạp file ZIP từ MinerU")
    zip_file = st.file_uploader("Kéo thả File .zip từ MinerU:", type=["zip"], key="offline_zip")
    
    if zip_file:
        base_name = zip_file.name.rsplit(".", 1)[0]
        json_data, images_dict = extract_zip_and_get_data(zip_file.getvalue())

        if json_data:
            st.success("✅ Đã nạp thành công dữ liệu ZIP!")
            st.markdown("---")

            col1, col2 = st.columns(2)

            with col1:
                try:
                    with st.spinner("Đang tạo Word Native Equation..."):
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