import os
import re
import uuid
from glob import glob
from datetime import datetime
import streamlit as st
from update_label_pdf import get_page_size, build_overlay, merge_and_write, read_n_text

os.makedirs('uploads', exist_ok=True)
os.makedirs('outputs', exist_ok=True)
os.makedirs('tmp', exist_ok=True)
os.environ['TMPDIR'] = os.path.abspath('tmp')

BATCH_OFFSET_X = -48.0
BATCH_OFFSET_Y = -5.0
BATCH_FONT_SIZE = 6

def infer_label_path(base_dir):
    env = os.environ.get('LABEL_PDF_PATH')
    if env and os.path.exists(env):
        return env
    fixed = os.path.join(base_dir, 'BQ20251210125617009000UY6.pdf')
    if os.path.exists(fixed):
        return fixed
    files = [p for p in glob(os.path.join(base_dir, 'BQ*.pdf')) if 'updated' not in os.path.basename(p).lower()]
    if not files:
        return None
    files.sort(key=lambda p: os.path.getmtime(p), reverse=True)
    return files[0]

def cleanup_dir(d, limit=10):
    try:
        files = [os.path.join(d, f) for f in os.listdir(d)]
        files = [p for p in files if os.path.isfile(p)]
        files.sort(key=lambda p: os.path.getmtime(p), reverse=True)
        for p in files[limit:]:
            try:
                os.remove(p)
            except Exception:
                pass
    except Exception:
        pass

def extract_digits_from_name(p):
    name = os.path.basename(p)
    base = os.path.splitext(name)[0]
    m = re.search(r"(?i)^skc_(\d+)", base)
    if m:
        return m.group(1)
    return ""

st.set_page_config(page_title='标签生成', layout='centered')
st.title('上传 SKC PDF 生成更新标签')

uploaded = st.file_uploader('拖拽或选择 SKC PDF 文件', type=['pdf'])

if uploaded is not None:
    name = uploaded.name.strip()
    if not re.fullmatch(r"(?i)skc_\d+(?:\s*\(\d+\))?\.pdf", name):
        st.error('文件名必须为 skc_[数字].pdf 格式，例如 skc_93244706336.pdf')
    else:
        base_dir = os.getcwd()
        label = infer_label_path(base_dir)
        if not label:
            st.error('未找到标签模板，请在当前目录放置 BQ*.pdf 或设置环境变量 LABEL_PDF_PATH')
        else:
            filename = f"{uuid.uuid4().hex}_{uploaded.name}"
            src_path = os.path.join('uploads', filename)
            with open(src_path, 'wb') as f:
                f.write(uploaded.read())
            cleanup_dir('uploads', 10)
            cleanup_dir('tmp', 10)
            skc_id = extract_digits_from_name(uploaded.name)
            n_text = read_n_text(skc_id)
            try:
                w, h = get_page_size(label)
                overlay = build_overlay(
                    label, w, h, src_path, n_text,
                    img_max_ratio=0.40,
                    img_lr_margin=6,
                    img_bottom_margin=0,
                    font_path=None,
                    font_bold_path=None,
                    batch_align='right',
                    batch_offset_x=BATCH_OFFSET_X,
                    batch_offset_y=BATCH_OFFSET_Y,
                    batch_font_size=int(BATCH_FONT_SIZE),
                    batch_font_weight='bold',
                    batch_length_align='left',
                    img_height_pt=0,
                    img_scale=1.0,
                    place='absolute',
                    abs_x=0,
                    abs_y=0,
                    render_dpi=240,
                )
                ts = datetime.now().strftime('%Y%m%d%H%M%S')
                out_name = f"标签_skc_{skc_id}_{ts}.pdf"
                out_path = os.path.join('outputs', out_name)
                merge_and_write(label, overlay, out_path)
                cleanup_dir('outputs', 10)
                with open(out_path, 'rb') as outf:
                    st.download_button('下载已更新的标签 PDF', data=outf.read(), file_name=out_name, mime='application/pdf')
                st.success(f'生成完成，使用模板: {label}')
            except Exception as e:
                st.error(f'处理失败: {e}')
