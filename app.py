import os
import io
import gc
import uuid
import copy
import shutil
import subprocess
from flask import Flask, request, jsonify, send_file, render_template
import numpy as np
import cv2
import fitz
from PIL import Image, ImageDraw, ImageOps
from pptx import Presentation
from pptx.util import Emu
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER

app = Flask(__name__, template_folder="templates/templates")
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
UPLOAD_ROOT = os.path.join(BASE_DIR, "workspace")
os.makedirs(UPLOAD_ROOT, exist_ok=True)


def get_session_dir(sid):
    path = os.path.join(UPLOAD_ROOT, sid)
    os.makedirs(path, exist_ok=True)
    return path


def pdf_to_pptx(pdf_path, pptx_out_path):
    doc = fitz.open(pdf_path)
    prs = Presentation()
    blank_layout = min(prs.slide_layouts, key=lambda l: len(l.placeholders))

    first_rect = doc[0].rect
    prs.slide_width = Emu(int(first_rect.width * 12700))
    prs.slide_height = Emu(int(first_rect.height * 12700))

    for page in doc:
        pix = page.get_pixmap(dpi=180)
        img_bytes = pix.tobytes("png")
        slide = prs.slides.add_slide(blank_layout)
        slide.shapes.add_picture(
            io.BytesIO(img_bytes), 0, 0, prs.slide_width, prs.slide_height
        )
    prs.save(pptx_out_path)


def render_pptx_previews(pptx_path, out_dir):
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if soffice:
        try:
            subprocess.run(
                [soffice, "--headless", "--convert-to", "pdf", "--outdir", out_dir, pptx_path],
                check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=60
            )
            pdf_name = os.path.splitext(os.path.basename(pptx_path))[0] + ".pdf"
            pdf_path = os.path.join(out_dir, pdf_name)
            if os.path.exists(pdf_path):
                doc = fitz.open(pdf_path)
                paths = []
                for idx, page in enumerate(doc):
                    pix = page.get_pixmap(dpi=130)
                    p = os.path.join(out_dir, f"slide_{idx}.png")
                    pix.save(p)
                    paths.append(p)
                return paths
        except Exception:
            pass

    prs = Presentation(pptx_path)
    sw, sh = prs.slide_width, prs.slide_height
    pw = 900
    ph = max(200, int(pw * (sh / sw)))
    paths = []

    for idx, slide in enumerate(prs.slides):
        canvas = Image.new("RGB", (pw, ph), (255, 255, 255))
        draw = ImageDraw.Draw(canvas)

        for shape in slide.shapes:
            x = int((shape.left / sw) * pw)
            y = int((shape.top / sh) * ph)
            w = max(1, int((shape.width / sw) * pw))
            h = max(1, int((shape.height / sh) * ph))

            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                try:
                    pic = Image.open(io.BytesIO(shape.image.blob)).convert("RGB")
                    pic = pic.resize((w, h), Image.Resampling.LANCZOS)
                    canvas.paste(pic, (x, y))
                except Exception:
                    draw.rectangle([x, y, x + w, y + h], outline=(0, 120, 255), width=2)
            else:
                draw.rectangle([x, y, x + w, y + h], outline=(180, 180, 180), width=1)
                if shape.has_text_frame and shape.text_frame.text.strip():
                    txt = shape.text_frame.text.strip()[:40]
                    draw.text((x + 4, y + 4), txt, fill=(40, 40, 40))

        p = os.path.join(out_dir, f"slide_{idx}.png")
        canvas.save(p)
        paths.append(p)
    return paths


def detect_photo_frames(slide, sw, sh):
    frames = []
    for idx, shape in enumerate(slide.shapes):
        is_pic = shape.shape_type == MSO_SHAPE_TYPE.PICTURE
        is_pic_ph = (
            shape.is_placeholder
            and shape.placeholder_format.type == PP_PLACEHOLDER.PICTURE
        )
        if is_pic or is_pic_ph:
            if (shape.width / sw) > 0.92 and (shape.height / sh) > 0.92:
                continue
            frames.append({
                "shape_index": idx,
                "x": round(shape.left / sw, 5),
                "y": round(shape.top / sh, 5),
                "w": round(shape.width / sw, 5),
                "h": round(shape.height / sh, 5),
            })
    return frames


def order_points(pts):
    rect = np.zeros((4, 2), dtype="float32")
    s = pts.sum(axis=1)
    rect[0] = pts[np.argmin(s)]
    rect[2] = pts[np.argmax(s)]
    diff = np.diff(pts, axis=1)
    rect[1] = pts[np.argmin(diff)]
    rect[3] = pts[np.argmax(diff)]
    return rect


def auto_perspective_crop(img_bgr):
    h, w = img_bgr.shape[:2]
    scale = 700.0 / max(h, w)
    small = cv2.resize(img_bgr, (int(w * scale), int(h * scale))) if scale < 1 else img_bgr.copy()

    gray = cv2.cvtColor(small, cv2.COLOR_BGR2GRAY)
    blur = cv2.GaussianBlur(gray, (5, 5), 0)
    edged = cv2.Canny(blur, 50, 150)
    edged = cv2.dilate(edged, np.ones((3, 3), np.uint8), iterations=1)

    cnts, _ = cv2.findContours(edged, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    cnts = sorted(cnts, key=cv2.contourArea, reverse=True)[:5]

    img_area = small.shape[0] * small.shape[1]
    for c in cnts:
        area = cv2.contourArea(c)
        if area < img_area * 0.20:
            continue
        peri = cv2.arcLength(c, True)
        approx = cv2.approxPolyDP(c, 0.02 * peri, True)
        if len(approx) == 4 and cv2.isContourConvex(approx):
            pts = approx.reshape(4, 2).astype("float32")
            if scale < 1:
                pts = pts / scale
            rect = order_points(pts)
            (tl, tr, br, bl) = rect
            maxW = int(max(np.linalg.norm(br - bl), np.linalg.norm(tr - tl)))
            maxH = int(max(np.linalg.norm(tr - br), np.linalg.norm(tl - bl)))
            if maxW > 50 and maxH > 50:
                dst = np.array([[0, 0], [maxW - 1, 0], [maxW - 1, maxH - 1], [0, maxH - 1]], dtype="float32")
                M = cv2.getPerspectiveTransform(rect, dst)
                return cv2.warpPerspective(img_bgr, M, (maxW, maxH))
    return img_bgr


def process_tile_photo(img_bytes, target_ratio, use_perspective=True, crop_mode="cover"):
    pil_img = Image.open(io.BytesIO(img_bytes))
    pil_img = ImageOps.exif_transpose(pil_img).convert("RGB")
    if max(pil_img.size) > 1600:
        pil_img.thumbnail((1600, 1600), Image.Resampling.LANCZOS)
    img_bgr = cv2.cvtColor(np.array(pil_img), cv2.COLOR_RGB2BGR)

    if use_perspective:
        img_bgr = auto_perspective_crop(img_bgr)

    h, w = img_bgr.shape[:2]
    current_ratio = w / h

    if crop_mode == "cover":
        if current_ratio > target_ratio:
            new_w = int(h * target_ratio)
            start_x = (w - new_w) // 2
            img_bgr = img_bgr[:, start_x:start_x + new_w]
        elif current_ratio < target_ratio:
            new_h = int(w / target_ratio)
            start_y = (h - new_h) // 2
            img_bgr = img_bgr[start_y:start_y + new_h, :]

    ok, buf = cv2.imencode(".jpg", img_bgr, [int(cv2.IMWRITE_JPEG_QUALITY), 90])
    return io.BytesIO(buf.tobytes())


def duplicate_slide(prs, template_slide):
    new_slide = prs.slides.add_slide(template_slide.slide_layout)
    for shape in list(new_slide.shapes):
        sp = shape._element
        sp.getparent().remove(sp)

    rId_map = {}
    for rId, rel in template_slide.part.rels.items():
        if "notesSlide" in rel.reltype or "slideLayout" in rel.reltype:
            continue
        if rel.is_external:
            rId_map[rId] = new_slide.part.relate_to(rel.target_ref, rel.reltype, is_external=True)
        else:
            rId_map[rId] = new_slide.part.relate_to(rel.target_part, rel.reltype)

    src_spTree = template_slide.shapes._spTree
    dst_spTree = new_slide.shapes._spTree
    for child in src_spTree:
        if child.tag.endswith("nvGrpSpPr") or child.tag.endswith("grpSpPr"):
            continue
        new_child = copy.deepcopy(child)
        for elem in new_child.iter():
            for k, v in list(elem.attrib.items()):
                if v in rId_map:
                    elem.attrib[k] = rId_map[v]
        dst_spTree.append(new_child)

    ns = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
    src_cSld = template_slide._element.cSld
    dst_cSld = new_slide._element.cSld
    src_bg = src_cSld.find(f"{ns}bg")
    if src_bg is not None:
        dst_bg = dst_cSld.find(f"{ns}bg")
        if dst_bg is not None:
            dst_cSld.remove(dst_bg)
        new_bg = copy.deepcopy(src_bg)
        for elem in new_bg.iter():
            for k, v in list(elem.attrib.items()):
                if v in rId_map:
                    elem.attrib[k] = rId_map[v]
        dst_cSld.insert(0, new_bg)

    return new_slide


@app.route("/")
def index():
    return render_template("templates/index.html")


@app.route("/upload_template", methods=["POST"])
def upload_template():
    file = request.files.get("template")
    if not file:
        return jsonify({"error": "File missing"}), 400

    for old_d in os.listdir(UPLOAD_ROOT):
        shutil.rmtree(os.path.join(UPLOAD_ROOT, old_d), ignore_errors=True)

    sid = str(uuid.uuid4())[:8]
    sdir = get_session_dir(sid)
    ext = os.path.splitext(file.filename)[1].lower()
    pptx_path = os.path.join(sdir, "template.pptx")

    if ext == ".pdf":
        pdf_path = os.path.join(sdir, "input.pdf")
        file.save(pdf_path)
        pdf_to_pptx(pdf_path, pptx_path)
    elif ext == ".pptx":
        file.save(pptx_path)
    else:
        return jsonify({"error": "Sirf .pptx ya .pdf upload karein"}), 400

    render_pptx_previews(pptx_path, sdir)
    prs = Presentation(pptx_path)
    slides_info = []
    for i, slide in enumerate(prs.slides):
        slides_info.append({
            "index": i,
            "preview_url": f"/preview/{sid}/{i}",
            "frames": detect_photo_frames(slide, prs.slide_width, prs.slide_height)
        })

    return jsonify({
        "session_id": sid,
        "slide_count": len(prs.slides),
        "slides": slides_info
    })


@app.route("/preview/<sid>/<int:idx>")
def get_preview(sid, idx):
    p = os.path.join(UPLOAD_ROOT, sid, f"slide_{idx}.png")
    return send_file(p, mimetype="image/png")


@app.route("/generate", methods=["POST"])
def generate_catalog():
    sid = request.form.get("session_id")
    template_idx = int(request.form.get("template_idx", 0))
    fx = float(request.form.get("x", 0.2))
    fy = float(request.form.get("y", 0.2))
    fw = float(request.form.get("w", 0.6))
    fh = float(request.form.get("h", 0.6))
    replace_shape_idx = request.form.get("replace_shape_idx", "")
    use_perspective = request.form.get("perspective", "true") == "true"
    crop_mode = request.form.get("crop_mode", "cover")

    photos = request.files.getlist("photos")
    if not sid or not photos:
        return jsonify({"error": "Session ya photos missing hain"}), 400

    sdir = get_session_dir(sid)
    pptx_path = os.path.join(sdir, "template.pptx")
    prs = Presentation(pptx_path)
    template_slide = prs.slides[template_idx]

    sw, sh = prs.slide_width, prs.slide_height
    left = Emu(int(fx * sw))
    top = Emu(int(fy * sh))
    width = Emu(int(fw * sw))
    height = Emu(int(fh * sh))
    target_ratio = (fw * sw) / max(1, (fh * sh))

    new_slide_elems = []
    for idx_p, photo in enumerate(photos):
        raw_bytes = photo.read()
        tile_name = os.path.splitext(photo.filename)[0]
        processed_stream = process_tile_photo(raw_bytes, target_ratio, use_perspective, crop_mode)
        del raw_bytes

        nslide = duplicate_slide(prs, template_slide)

        for shp in nslide.shapes:
            if shp.has_text_frame:
                for para in shp.text_frame.paragraphs:
                    for run in para.runs:
                        if "{{NAME}}" in run.text:
                            run.text = run.text.replace("{{NAME}}", tile_name)

        spTree = nslide.shapes._spTree
        insert_pos = None
        if replace_shape_idx != "":
            ridx = int(replace_shape_idx)
            if 0 <= ridx < len(nslide.shapes):
                old_el = nslide.shapes[ridx]._element
                insert_pos = list(spTree).index(old_el)
                spTree.remove(old_el)

        pic_shape = nslide.shapes.add_picture(processed_stream, left, top, width, height)
        if insert_pos is not None:
            pic_el = pic_shape._element
            spTree.remove(pic_el)
            spTree.insert(insert_pos, pic_el)

        del processed_stream
        if idx_p % 10 == 0:
            gc.collect()

        new_slide_elems.append(prs.slides._sldIdLst[-1])

    sldIdLst = prs.slides._sldIdLst
    for el in new_slide_elems:
        sldIdLst.remove(el)
    for offset, el in enumerate(new_slide_elems):
        sldIdLst.insert(template_idx + 1 + offset, el)

    old_rId = sldIdLst[template_idx].rId
    prs.part.drop_rel(old_rId)
    del sldIdLst[template_idx]

    out_path = os.path.join(sdir, "Tiles_Catalog_Final.pptx")
    prs.save(out_path)
    return send_file(out_path, as_attachment=True, download_name="Tiles_Catalog_Final.pptx")


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
