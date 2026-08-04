import io
import json
import os
import re
import tempfile
import zipfile
import streamlit as st
from bs4 import BeautifulSoup

# Import tkinter để mở cửa sổ chọn thư mục trên máy
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
# HÀM MỞ CỬA SỔ CHỌN FOLDER NATIVE
# ==========================================
def select_folder():
    if tk:
        root = tk.Tk()
        root.withdraw()
        root.wm_attributes("-topmost", 1)
        folder_selected = filedialog.askdirectory(master=root)
        root.destroy()
        return folder_selected
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

def get_image_bytes(img_filename, images_dict, local_img_dir=None, uploaded_images_map=None):
    if not img_filename:
        return None
    clean_name = os.path.basename(img_filename)
    
    # Ưu tiên 1: Lấy từ images_dict (nếu đọc file ZIP)
    if clean_name in images_dict:
        return io.BytesIO(images_dict[clean_name])
    
    # Ưu tiên 2: Lấy từ các file ảnh được upload trực tiếp
    if uploaded_images_map and clean_name in uploaded_images_map:
        return io.BytesIO(uploaded_images_map[clean_name])

    # Ưu tiên 3: Tự động đọc trực tiếp từ thư mục trên máy tính
    if local_img_dir and os.path.exists(local_img_dir):
        local_path = os.path.join(local_img_dir, clean_name)
        if os.path.exists(local_path):
            with open(local_path, "rb") as f:
                return io.BytesIO(f.read())
    return None

def process_html_table_md(table_html, images_dict, temp_dir=None, local_img_dir=None, uploaded_images_map=None):
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
                        img_stream = get_image_bytes(img_src, images_dict, local_img_dir, uploaded_images_map)
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

def convert_json_to_markdown(json_data, images_dict, temp_dir=None, local_img_dir=None, uploaded_images_map=None):
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
                        img_stream = get_image_bytes(img_path, images_dict, local_img_dir, uploaded_images_map)
                        if img_stream and temp_dir:
                            img_counter += 1
                            local_img_path = os.path.join(temp_dir, f"img_{img_counter}.png")
                            with open(local_img_path, "wb") as f:
                                f.write(img_stream.getvalue())
                            md_lines.append(f"![Hình ảnh]({local_img_path})\n\n")

                elif b_type == "table":
                    table_html = item.get("table_html", item.get("html", ""))
                    if table_html:
                        md_lines.append(process_html_table_md(table_html, images_dict, temp_dir, local_img_dir, uploaded_images_map))
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
                                img_stream = get_image_bytes(img_path, images_dict, local_img_dir, uploaded_images_map)
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
                                md_lines.append(process_html_table_md(table_html, images_dict, temp_dir, local_img_dir, uploaded_images_map))

    return "".join(md_lines)

def convert_json_to_docx_pandoc_bytes(json_data, images_dict, local_img_dir=None, uploaded_images_map=None):
    with tempfile.TemporaryDirectory() as temp_dir:
        md_text = convert_json_to_markdown(json_data, images_dict, temp_dir=temp_dir, local_img_dir=local_img_dir, uploaded_images_map=uploaded_images_map)

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
st.set_page_config(page_title="MinerU Local Converter", page_icon="📦", layout="wide")
st.title("📦 MinerU Local Auto-Save & Converter")

st.subheader("Chọn Phương Thức Nạp Dữ Liệu")
input_mode = st.radio("Chế độ:", ["📦 Chọn File ZIP", "📄 Chọn File JSON (Kèm Ảnh)"], horizontal=True)

json_data = {}
images_dict = {}
local_img_dir = None
uploaded_images_map = {}
base_name = "Converted_Doc"

if input_mode == "📦 Chọn File ZIP":
    zip_file = st.file_uploader("Kéo thả File .zip từ MinerU:", type=["zip"], key="offline_zip")
    if zip_file:
        base_name = zip_file.name.rsplit(".", 1)[0]
        json_data, images_dict = extract_zip_and_get_data(zip_file.getvalue())

else:
    json_file = st.file_uploader("1. Tải lên file JSON (ví dụ: layout.json):", type=["json"], key="offline_json")
    
    st.markdown("---")
    st.markdown("**2. Cung cấp nguồn ảnh (Chọn một trong hai cách dưới đây):**")
    
    img_source_type = st.radio("Nguồn ảnh:", ["📁 Chọn thư mục hoặc nhập đường dẫn thư mục ảnh trên máy", "📤 Tải lên các file ảnh thủ công"], horizontal=True)
    
    if img_source_type == "📁 Chọn thư mục hoặc nhập đường dẫn thư mục ảnh trên máy":
        col_path, col_btn_path = st.columns([3, 1])
        
        if "local_img_path" not in st.session_state:
            st.session_state.local_img_path = ""

        with col_path:
            local_img_dir = st.text_input("Đường dẫn thư mục images:", value=st.session_state.local_img_path, placeholder="Ví dụ: D:\\Downloads\\images")
            st.session_state.local_img_path = local_img_dir

        with col_btn_path:
            st.write("") 
            st.write("")
            if st.button("📁 Chọn Thư Mục"):
                chosen_img_dir = select_folder()
                if chosen_img_dir:
                    st.session_state.local_img_path = chosen_img_dir
                    st.rerun()
    else:
        uploaded_image_files = st.file_uploader(
            "Tải lên các file ảnh (PNG, JPG, JPEG):", 
            type=["png", "jpg", "jpeg"], 
            accept_multiple_files=True,
            key="multi_img_upload"
        )
        if uploaded_image_files:
            for img_file in uploaded_image_files:
                uploaded_images_map[img_file.name] = img_file.getvalue()

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
                docx_pandoc = convert_json_to_docx_pandoc_bytes(
                    json_data, 
                    images_dict, 
                    local_img_dir=local_img_dir, 
                    uploaded_images_map=uploaded_images_map
                )
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
        md_content = convert_json_to_markdown(
            json_data, 
            images_dict, 
            local_img_dir=local_img_dir, 
            uploaded_images_map=uploaded_images_map
        )
        st.download_button(
            label="📄 Tải File Markdown (.md)",
            data=md_content,
            file_name=f"{base_name}.md",
            mime="text/markdown",
            use_container_width=True,
        )