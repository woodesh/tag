import argparse
import os
import re
import io
import pdfplumber
from reportlab.pdfgen import canvas
from reportlab.lib.colors import white, black
from reportlab.lib.utils import ImageReader
from PyPDF2 import PdfReader, PdfWriter
import pypdfium2 as pdfium
from PIL import Image
from PIL import ImageStat
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

def render_source_pdf_to_image(source_path, dpi=200):
    doc = pdfium.PdfDocument(source_path)
    page = doc.get_page(0)
    bmp = page.render(scale=dpi / 72.0)
    img = bmp.to_pil()
    bmp.close()
    page.close()
    doc.close()
    return img

def read_n_text(n_value):
    p = str(n_value)
    if os.path.exists(p):
        with open(p, "r", encoding="utf-8") as f:
            return f.read().strip()
    return p.strip()

def get_page_size(label_pdf_path):
    reader = PdfReader(label_pdf_path)
    page = reader.pages[0]
    mb = page.mediabox
    w = float(mb.right) - float(mb.left)
    h = float(mb.top) - float(mb.bottom)
    return w, h

def find_batch_area(label_pdf_path):
    with pdfplumber.open(label_pdf_path) as pdf:
        page = pdf.pages[0]
        words = page.extract_words() or []
        target_idx = None
        for i in range(len(words) - 1):
            t1 = words[i]["text"].lower()
            t2 = words[i + 1]["text"].lower()
            if t1 == "batch" and t2.startswith("number"):
                target_idx = i
                break
        if target_idx is None:
            for i in range(len(words)):
                if "batch" in words[i]["text"].lower():
                    target_idx = i
                    break
        if target_idx is not None:
            x1 = words[target_idx]["x1"]
            top = words[target_idx].get("top", 0)
            bottom = words[target_idx].get("bottom", top + 12)
            return x1 + 2, top, 140, bottom - top
    return None

def compute_free_bottom_height(label_pdf_path, page_height):
    dpi = 300
    doc = pdfium.PdfDocument(label_pdf_path)
    page = doc.get_page(0)
    bmp = page.render(scale=dpi / 72.0)
    img = bmp.to_pil().convert("L")
    w, h = img.size
    threshold = 250
    margin_px = 0
    for y in range(h - 1, -1, -1):
        row = img.crop((0, y, w, y + 1))
        mean = ImageStat.Stat(row).mean[0]
        if mean >= threshold:
            margin_px += 1
        else:
            break
    bmp.close()
    page.close()
    doc.close()
    return margin_px * 72.0 / dpi

def find_barcode_area(label_pdf_path, page_width, page_height, left_margin=6, right_margin=6, top_pad=2, bottom_pad=2):
    with pdfplumber.open(label_pdf_path) as pdf:
        page = pdf.pages[0]
        words = page.extract_words() or []
        def is_digits(w):
            t = w.get("text", "")
            return t.isdigit() and len(t) >= 8
        digit_words = [w for w in words if is_digits(w)]
        if not digit_words:
            return None
        top_digit = min(digit_words, key=lambda w: w.get("top", 0))
        bottom_digit = max(digit_words, key=lambda w: w.get("bottom", 0))
        top_y = top_digit.get("bottom", top_digit.get("top", 0)) + top_pad
        bottom_y = bottom_digit.get("top", bottom_digit.get("bottom", 0)) - bottom_pad
        if bottom_y <= top_y:
            return None
        x0 = left_margin
        x1 = page_width - right_margin
        # avoid right-side non-digit words inside band (e.g., PVC)
        band_words = [w for w in words if w.get("top", 0) >= top_y and w.get("bottom", 0) <= bottom_y]
        non_digits = [w for w in band_words if not is_digits(w)]
        if non_digits:
            min_right_non_digit_x0 = min(w.get("x0", page_width) for w in non_digits)
            if min_right_non_digit_x0 < x1:
                x1 = max(x0 + 10, min_right_non_digit_x0 - 4)
        w = max(8, x1 - x0)
        h = bottom_y - top_y
        y = page_height - top_y - h
        return x0, y, w, h

def try_register_font(font_path=None, font_bold_path=None):
    reg_candidates = []
    bold_candidates = []
    if font_path:
        reg_candidates.append(font_path)
    if font_bold_path:
        bold_candidates.append(font_bold_path)
    bases = ["/Library/Fonts", "/System/Library/Fonts", os.path.expanduser("~/Library/Fonts")]
    for base in bases:
        for name in [
            "SourceHanSansCN-Regular.otf",
            "Source Han Sans CN Regular.otf",
            "SourceHanSansSC-Regular.otf",
            "Source Han Sans SC Regular.otf",
            "SourceHanSans-Regular.otf",
        ]:
            reg_candidates.append(os.path.join(base, name))
        for name in [
            "SourceHanSansCN-Bold.otf",
            "Source Han Sans CN Bold.otf",
            "SourceHanSansSC-Bold.otf",
            "Source Han Sans SC Bold.otf",
            "SourceHanSans-Bold.otf",
        ]:
            bold_candidates.append(os.path.join(base, name))
    reg_name = "Helvetica"
    bold_name = "Helvetica-Bold"
    for p in reg_candidates:
        if os.path.exists(p):
            try:
                pdfmetrics.registerFont(TTFont("SourceHanSansCN-Regular", p))
                reg_name = "SourceHanSansCN-Regular"
                break
            except Exception:
                pass
    for p in bold_candidates:
        if os.path.exists(p):
            try:
                pdfmetrics.registerFont(TTFont("SourceHanSansCN-Bold", p))
                bold_name = "SourceHanSansCN-Bold"
                break
            except Exception:
                pass
    return reg_name, bold_name

def build_overlay(label_pdf_path, width, height, source_pdf_path, batch_text, img_max_ratio=0.40, img_lr_margin=6, img_bottom_margin=0, font_path=None, font_bold_path=None, batch_align="right", batch_offset_x=0, batch_offset_y=0, batch_font_size=12, batch_font_weight="regular", batch_length_align="none", img_height_pt=0, img_scale=1.0, place="barcode", abs_x=0, abs_y=0, render_dpi=200):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(width, height))
    img = render_source_pdf_to_image(source_pdf_path, dpi=render_dpi)
    iw, ih = img.size
    free_bottom_h = compute_free_bottom_height(label_pdf_path, height)
    max_h = min(height * img_max_ratio, max(18, free_bottom_h - img_bottom_margin - 4))
    target_w = max(12, width - 2 * img_lr_margin)
    if img_height_pt and img_height_pt > 0:
        s = img_height_pt / ih
    else:
        s = min(target_w / iw, max_h / ih)
        s *= max(0.1, img_scale)
        s = min(s, target_w / iw, max_h / ih)
    dw = iw * s
    dh = ih * s
    if place == "absolute":
        dw = iw * 72.0 / render_dpi
        dh = ih * 72.0 / render_dpi
        x_img = abs_x
        y_img = abs_y
        c.drawImage(ImageReader(img), x_img, y_img, width=dw, height=dh, preserveAspectRatio=True)
    elif place == "barcode":
        area_bc = find_barcode_area(label_pdf_path, width, height)
        if area_bc:
            bx, by, bw, bh = area_bc
            c.setFillColor(white)
            c.rect(bx, by, bw, bh, fill=1, stroke=0)
            s_cover = min(bw / iw, bh / ih)
            dw2 = iw * s_cover
            dh2 = ih * s_cover
            x_img = bx + (bw - dw2) / 2
            y_img = by + (bh - dh2) / 2
            c.drawImage(ImageReader(img), x_img, y_img, width=dw2, height=dh2, preserveAspectRatio=True)
        else:
            x_img = (width - dw) / 2
            y_img = img_bottom_margin
            c.drawImage(ImageReader(img), x_img, y_img, width=dw, height=dh, preserveAspectRatio=True)
    else:
        x_img = (width - dw) / 2
        y_img = img_bottom_margin
        c.drawImage(ImageReader(img), x_img, y_img, width=dw, height=dh, preserveAspectRatio=True)
    font_reg, font_bold = try_register_font(font_path, font_bold_path)
    font_name = font_bold if batch_font_weight == "bold" else font_reg
    area = find_batch_area(label_pdf_path)
    if area:
        x, top, w, h = area
        y = height - (top + h)
        c.setFont(font_name, batch_font_size)
        tw = pdfmetrics.stringWidth(batch_text, font_name, batch_font_size)
        tw_base = pdfmetrics.stringWidth("93244706336", font_name, batch_font_size)
        if batch_length_align == "left":
            length_adjust = (tw - tw_base)
        elif batch_length_align == "center":
            length_adjust = (tw - tw_base) / 2.0
        else:
            length_adjust = 0.0
        if batch_align == "right":
            x_draw = x + w - tw - 1 + batch_offset_x + length_adjust
        else:
            x_draw = x + 1 + batch_offset_x
        y_draw = y + 2 + batch_offset_y
        c.setFillColor(white)
        c.rect(x_draw - 1, y_draw - 2, tw + 2, batch_font_size + 4, fill=1, stroke=0)
        c.setFillColor(black)
        c.drawString(x_draw, y_draw, batch_text)
    else:
        pass
    c.save()
    buf.seek(0)
    return buf

def merge_and_write(label_pdf_path, overlay_buf, output_path):
    base_reader = PdfReader(label_pdf_path)
    overlay_reader = PdfReader(overlay_buf)
    writer = PdfWriter()
    base_page = base_reader.pages[0]
    overlay_page = overlay_reader.pages[0]
    try:
        base_page.merge_page(overlay_page)
    except Exception:
        base_page.mergePage(overlay_page)
    writer.add_page(base_page)
    for i in range(1, len(base_reader.pages)):
        writer.add_page(base_reader.pages[i])
    with open(output_path, "wb") as f:
        writer.write(f)

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--label", required=True)
    parser.add_argument("--source", required=True)
    parser.add_argument("--n", required=True)
    parser.add_argument("--output", required=False)
    parser.add_argument("--img_max_ratio", type=float, default=0.40)
    parser.add_argument("--img_lr_margin", type=float, default=6)
    parser.add_argument("--img_bottom_margin", type=float, default=0)
    parser.add_argument("--font", required=False)
    parser.add_argument("--font_bold", required=False)
    parser.add_argument("--batch_align", choices=["left", "right"], default="right")
    parser.add_argument("--batch_offset_x", type=float, default=0)
    parser.add_argument("--batch_offset_y", type=float, default=0)
    parser.add_argument("--batch_font_size", type=int, default=12)
    parser.add_argument("--batch_font_weight", choices=["regular","bold"], default="regular")
    parser.add_argument("--batch_length_align", choices=["none","left","center"], default="none")
    parser.add_argument("--img_height_pt", type=float, default=0)
    parser.add_argument("--img_scale", type=float, default=1.0)
    parser.add_argument("--place", choices=["bottom","barcode","absolute"], default="barcode")
    parser.add_argument("--abs_x", type=float, default=0)
    parser.add_argument("--abs_y", type=float, default=0)
    parser.add_argument("--render_dpi", type=int, default=200)
    args = parser.parse_args()
    w, h = get_page_size(args.label)
    batch_text = read_n_text(args.n)
    overlay = build_overlay(
        args.label, w, h, args.source, batch_text,
        img_max_ratio=args.img_max_ratio,
        img_lr_margin=args.img_lr_margin,
        img_bottom_margin=args.img_bottom_margin,
        font_path=args.font,
        font_bold_path=args.font_bold,
        batch_align=args.batch_align,
        batch_offset_x=args.batch_offset_x,
        batch_offset_y=args.batch_offset_y,
        batch_font_size=args.batch_font_size,
        batch_font_weight=args.batch_font_weight,
        batch_length_align=args.batch_length_align,
        img_height_pt=args.img_height_pt,
        img_scale=args.img_scale,
        place=args.place,
        abs_x=args.abs_x,
        abs_y=args.abs_y,
        render_dpi=args.render_dpi,
    )
    out = args.output or os.path.join(os.path.dirname(args.label), "label_updated.pdf")
    merge_and_write(args.label, overlay, out)

if __name__ == "__main__":
    main()
