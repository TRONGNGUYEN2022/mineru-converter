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

# Kiểm tra thư viện tkinter (chỉ hoạt động khi chạy Local)
try:
    import tkinter as tk
    from tkinter import filedialog
except ImportError:
    tk = None

# Nạp các thư viện phụ thuộc
try:
    from docx import Document
except ImportError:
    os.system("pip install python-docx")
    from docx import Document

try:
    import pypandoc
except ImportError:
    os.system("pip install pypandoc")
    import pypandoc

# ==========================================
# 1. CẤU HÌNH HỆ THỐNG VÀ THƯ MỤC LƯU
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

def select_folder():
    """Hàm mở cửa sổ chọn folder (Chỉ dùng được khi chạy Local)"""
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
# 2. XỬ LÝ DỮ LIỆU LAYOUT JSON & LATEX
# ==========================================
def format_latex_string(latex_str):
    """Chuẩn hóa chuỗi LaTeX để Pandoc đọc công thức toán chính xác"""
    if not latex_str:
        return ""
    latex_clean = str(latex_str).strip()
    if latex_clean.startswith("$") and latex_clean.endswith("$"):
        latex_clean = latex_clean[1:-1].strip()
    latex_clean = latex_clean.replace("\n", " ")
    return f"${latex_clean}$"

def get_image_bytes(img_filename, images_dict, local_img_dir=None):
    """Đọc ảnh từ ZIP hoặc từ thư mục máy tính"""
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
    """Chuyển đổi bảng HTML trong MinerU sang Markdown Table"""
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
    """Đọc dữ liệu MinerU JSON và chuyển thành chuỗi Markdown chuẩn"""
    md_lines = []
    img_counter = 0

    if not json_data:
        return ""

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
    """Dùng Pandoc để render Markdown thành Word (.docx) Native Math Equation"""
    with tempfile.TemporaryDirectory() as temp_dir:
        md_text = convert_json_to_markdown(json_data, images_dict, temp_dir=temp_dir, local_img_dir=local_img_dir)

        if not md_text.strip():
            raise ValueError("Nội dung bóc tách bị rỗng!")

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
    """Bóc tách file ZIP trả về từ MinerU"""
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

    if not selected_file and json_files:
        selected_file = list(json_files.keys())[0]

    json_data = json_files.get(selected_file, {}) if selected_file else {}
    return json_data, images_dict

# ==========================================
# 3. GIAO DIỆN CHÍNH (STREAMLIT UI)
# ==========================================
st.set_page_config(page_title="MinerU Converter", page_icon="⚡", layout="wide")
st.title("⚡ MinerU Extractor & Word Native Converter")

# --- SIDEBAR CẤU HÌNH ---
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

    st.markdown("---")
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
            else:
                st.toast("⚠️ Chọn thư mục chỉ hoạt động khi chạy app Local trên máy tính!")
    with col_dir2:
        if st.button("🔄 Mặc định", use_container_width=True):
            st.session_state.output_dir = DEFAULT_OUTPUT_DIR
            st.rerun()

# --- TABS NỘI DUNG ---
tab1, tab2 = st.tabs([
    "🚀 Trích xuất API Trực Tiếp",
    "📦 Convert File Sẵn Có (ZIP / JSON)",
])

# --- TAB 1: TRÍCH XUẤT API ONLINE / LOCAL ---
with tab1:
    uploaded_file = st.file_uploader(
        "Tải lên file tài liệu (PDF, PNG, JPG, DOCX...):",
        type=["pdf", "docx", "doc", "pptx", "png", "jpg", "jpeg"],
        key="api_file",
    )

    if uploaded_file and st.button("🚀 Bắt đầu trích xuất", type="primary", use_container_width=True):
        token = st.session_state.api_key.strip()
        headers = {"Authorization": f"Bearer {token}"}
        base_name = uploaded_file.name.rsplit(".", 1)[0]
        status_box = st.empty()
        progress_bar = st.progress(0)

        try:
            # 1. Xin cấp URL Upload
            status_box.info("🔑 [1/3] Đang gửi yêu cầu tới MinerU API...")
            res_url = requests.post(
                "https://mineru.net/api/v4/file-urls/batch",
                headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
                json={"files": [{"name": uploaded_file.name}]},
                timeout=20
            )
            url_json = res_url.json()
            if res_url.status_code != 200 or url_json.get("code") != 0:
                status_box.error(f"❌ MinerU từ chối: {url_json.get('msg', 'API Key không hợp lệ hoặc hết lượt dùng')}")
                st.stop()

            batch_id = url_json["data"]["batch_id"]
            upload_url = url_json["data"]["file_urls"][0]
            progress_bar.progress(30)

            # 2. Upload file
            status_box.info("📤 [2/3] Đang tải file lên MinerU Storage...")
            put_res = requests.put(upload_url, data=uploaded_file.getvalue(), timeout=120)
            if put_res.status_code not in [200, 201]:
                status_box.error(f"❌ Upload file thất bại (HTTP {put_res.status_code})")
                st.stop()
            progress_bar.progress(50)

            # 3. Kích hoạt Task Phân tích
            status_box.info("⚡ [3/3] Đang kích hoạt tiến trình bóc tách công thức toán...")
            task_res = requests.post(
                "https://mineru.net/api/v4/extract/task",
                headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
                json={"batch_id": batch_id, "is_ocr": True},
                timeout=20
            )
            
            # Polling kiểm tra kết quả
            query_url = f"https://mineru.net/api/v4/extract/task/batch?batch_id={batch_id}"
            start_time = time.time()
            zip_url = None

            while True:
                elapsed = int(time.time() - start_time)
                if elapsed > 180:
                    status_box.error("⏱️ Quá thời gian chờ (3 phút)! Vui lòng thử lại với file nhỏ hơn hoặc kiểm tra credits tài khoản.")
                    st.stop()

                status_res = requests.get(query_url, headers=headers, timeout=10)
                if status_res.status_code == 200:
                    res_data = status_res.json().get("data", {})
                    extract_list = res_data.get("extract_result") or res_data.get("extract_results") or []

                    if extract_list:
                        item = extract_list[0]
                        state = str(item.get("state") or item.get("extract_state") or item.get("status") or "").lower()

                        status_box.info(f"⚡ Trạng thái MinerU: **{state.upper()}** ({elapsed}s)...")

                        if state in ["done", "success", "finished"]:
                            zip_url = item.get("full_zip_url") or item.get("download_url") or item.get("zip_url")
                            progress_bar.progress(100)
                            break
                        elif state in ["failed", "error"]:
                            status_box.error(f"❌ MinerU xử lý thất bại: {item.get('err_msg', 'Lỗi không xác định')}")
                            st.stop()

                time.sleep(2)

            # 4. Tải ZIP & Tạo Word
            if zip_url:
                status_box.info("⚡ Đang tự động tạo file Word Native Equation...")
                zip_res = requests.get(zip_url)
                
                if zip_res.status_code == 200:
                    zip_bytes = zip_res.content
                    
                    os.makedirs(st.session_state.output_dir, exist_ok=True)
                    saved_zip_path = os.path.join(st.session_state.output_dir, f"{base_name}_mineru.zip")
                    
                    try:
                        with open(saved_zip_path, "wb") as f:
                            f.write(zip_bytes)
                    except Exception:
                        pass # Bỏ qua nếu môi trường Cloud không cho ghi ổ cứng local

                    json_data, images_dict = extract_zip_and_get_data(zip_bytes)

                    if json_data:
                        status_box.success(f"🎉 Hoàn tất trích xuất trong **{int(time.time() - start_time)} giây**!")
                        st.markdown("---")
                        
                        col1, col2 = st.columns(2)
                        with col1:
                            try:
                                with st.spinner("Đang dựng file Word (Native Math Equation)..."):
                                    docx_pandoc = convert_json_to_docx_pandoc_bytes(json_data, images_dict)
                                st.download_button(
                                    label="📐 Tải File Word (.docx)",
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
        except Exception as e:
            status_box.error(f"❌ Lỗi kết nối: {e}")

# --- TAB 2: CONVERT FILE SẴN CÓ ---
with tab2:
    st.subheader("Chọn Phương Thức Nạp Dữ Liệu")
    input_mode = st.radio("Chế độ:", ["📦 Chọn File ZIP", "📄 Chọn File JSON (Tự đọc ảnh từ Folder)"], horizontal=True)

    json_data = {}
    images_dict = {}
    local_img_dir = None
    base_name = "Converted_Doc"

    if input_mode == "📦 Chọn File ZIP":
        zip_file = st.file_uploader("Kéo thả File .zip tải về từ MinerU:", type=["zip"], key="offline_zip")
        if zip_file:
            base_name = zip_file.name.rsplit(".", 1)[0]
            json_data, images_dict = extract_zip_and_get_data(zip_file.getvalue())

    else:
        json_file = st.file_uploader("1. Tải lên file JSON (layout.json):", type=["json"], key="offline_json")
        st.markdown("**2. Đường dẫn thư mục chứa Ảnh trên máy tính:**")
        col_path, col_btn_path = st.columns([3, 1])
        
        if "local_img_path" not in st.session_state:
            st.session_state.local_img_path = ""

        with col_path:
            local_img_dir = st.text_input("Đường dẫn thư mục images:", value=st.session_state.local_img_path, placeholder="C:\\downloads\\images")
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
                    label="📐 Tải File Word (.docx)",
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