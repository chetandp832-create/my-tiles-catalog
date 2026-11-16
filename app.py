import os
import io
import gc
import json
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

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
TPL_DIR = os.path.join(BASE_DIR, "templates", "templates")
if not os.path.exists(os.path.join(TPL_DIR, "index.html")):
    TPL_DIR = os.path.join(BASE_DIR, "templates")

app = Flask(__name__, template_folder=TPL_DIR)
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


def walk_shapes_with_bounds(shapes, off_x=0.0, off_y=0.0, scale_x=1.0, scale_y=1.0, is_top=True):
    results = []
    for shape in shapes:
        ax = off_x + shape.left * scale_x
        ay = off_y + shape.top * scale_y
        aw = shape.width * scale_x
        ah = shape.height * scale_y

        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            try:
                xfrm = shape._element.grpSpPr.xfrm
                ch_ox = float(xfrm.chOff.x) if xfrm.chOff is not None else 0.0
                ch_oy = float(xfrm.chOff.y) if xfrm.chOff is not None else 0.0
                ch_cx = float(xfrm.chExt.cx) if xfrm.chExt is not None and xfrm.chExt.cx else 1.0
                ch_cy = float(xfrm.chExt.cy) if xfrm.chExt is not None and xfrm.chExt.cy else 1.0
                sx = aw / ch_cx
                sy = ah / ch_cy
                results.extend(
                    walk_shapes_with_bounds(
                        shape.shapes,
                        off_x=ax - ch_ox * sx,
                        off_y=ay - ch_oy * sy,
                        scale_x=sx,
                        scale_y=sy,
                        is_top=False,
                    )
                )
            except Exception:
                pass
        else:
            results.append((shape, ax, ay, aw, ah, is_top))
    return results


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

        for shape, ax, ay, aw, ah, _ in walk_shapes_with_bounds(slide.shapes):
            x = int((ax / sw) * pw)
            y = int((ay / sh) * ph)
            w = max(1, int((aw / sw) * pw))
            h = max(1, int((ah / sh) * ph))

            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                try:
                    pic = Image.open(io.BytesIO(shape.image.blob)).convert("RGB")
                    pic = pic.resize((w, h), Image.Resampling.LANCZOS)
                    canvas.paste(pic, (x, y))
                except Exception:
                    draw.rectangle([x, y, x + w, y + h], outline=(0, 120, 255), width=2)
            else:
                draw.rectangle([x, y, x + w, y + h], outline=(190, 190, 190), width=1)
                if shape.has_text_frame and shape.text_frame.text.strip():
                    txt = shape.text_frame.text.strip()[:35]
                    draw.text((x + 4, y + 4), txt, fill=(40, 40, 40))

        p = os.path.join(out_dir, f"slide_{idx}.png")
        canvas.save(p)
        paths.append(p)
    return paths


def is_photo_placeholder_text(txt):
    t = txt.lower().strip()
    return ("add" in t and ("photo" in t or "tile" in t or "image" in t)) or t in ("photo", "tile photo", "image")


def detect_photo_frames(slide, sw, sh):
    frames = []
    for shape, ax, ay, aw, ah, _ in walk_shapes_with_bounds(slide.shapes):
        wn, hn = aw / sw, ah / sh
        if wn > 0.90 and hn > 0.90:
            continue
        if wn < 0.08 or hn < 0.04:
            continue

        is_pic = shape.shape_type == MSO_SHAPE_TYPE.PICTURE
        is_pic_ph = (
            shape.is_placeholder
            and shape.placeholder_format.type == PP_PLACEHOLDER.PICTURE
        )
        is_txt_frame = shape.has_text_frame and is_photo_placeholder_text(shape.text_frame.text)

        if is_pic or is_pic_ph or is_txt_frame:
            frames.append({
                "x": round(ax / sw, 5),
                "y": round(ay / sh, 5),
                "w": round(wn, 5),
                "h": round(hn, 5),
            })

    frames.sort(key=lambda f: (round(f["y"], 2), f["x"]))
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
    if max(pil_img.size) > 1500:
        pil_img.thumbnail((1500, 1500), Image.Resampling.LANCZOS)
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


def overlap_fraction(ax, ay, aw, ah, bx, by, bw, bh):
    ix1, iy1 = max(ax, bx), max(ay, by)
    ix2, iy2 = min(ax + aw, bx + bw), min(ay + ah, by + bh)
    if ix2 <= ix1 or iy2 <= iy1:
        return 0.0
    inter = (ix2 - ix1) * (iy2 - iy1)
    min_area = max(1e-9, min(aw * ah, bw * bh))
    return inter / min_area


def place_photo_in_box(slide, sw, sh, box, img_stream):
    fx, fy, fw, fh = box["x"], box["y"], box["w"], box["h"]
    spTree = slide.shapes._spTree
    insert_pos = None

    for shape, ax, ay, aw, ah, is_top in walk_shapes_with_bounds(slide.shapes):
        nx, ny, nw, nh = ax / sw, ay / sh, aw / sw, ah / sh
        if nw > 0.90 and nh > 0.90:
            continue
        if overlap_fraction(nx, ny, nw, nh, fx, fy, fw, fh) > 0.55:
            is_pic = shape.shape_type == MSO_SHAPE_TYPE.PICTURE
            is_pic_ph = shape.is_placeholder and shape.placeholder_format.type == PP_PLACEHOLDER.PICTURE
            if (is_pic or is_pic_ph) and is_top:
                el = shape._element
                if el.getparent() == spTree:
                    insert_pos = list(spTree).index(el)
                    spTree.remove(el)
            elif shape.has_text_frame and is_photo_placeholder_text(shape.text_frame.text):
                for para in shape.text_frame.paragraphs:
                    for run in para.runs:
                        run.text = ""

    left = Emu(int(fx * sw))
    top = Emu(int(fy * sh))
    width = Emu(int(fw * sw))
    height = Emu(int(fh * sh))
    pic_shape = slide.shapes.add_picture(img_stream, left, top, width, height)
    if insert_pos is not None:
        pic_el = pic_shape._element
        spTree.remove(pic_el)
        spTree.insert(insert_pos, pic_el)


def resolve_slide_boxes(slide, sw, sh, fallback_boxes, deltas, auto_frame_fit):
    if not auto_frame_fit:
        return fallback_boxes
    detected = detect_photo_frames(slide, sw, sh)
    if len(detected) == len(fallback_boxes) and len(detected) > 0:
        out = []
        for i, d in enumerate(detected):
            dl = deltas[i] if i < len(deltas) else {"dx": 0, "dy": 0, "dw": 0, "dh": 0}
            out.append({
                "x": max(0.0, min(0.95, d["x"] + dl.get("dx", 0))),
                "y": max(0.0, min(0.95, d["y"] + dl.get("dy", 0))),
                "w": max(0.03, min(1.0, d["w"] + dl.get("dw", 0))),
                "h": max(0.03, min(1.0, d["h"] + dl.get("dh", 0))),
            })
        return out
    return fallback_boxes


@app.route("/")
def index():
    return render_template("index.html")


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
    fill_mode = request.form.get("fill_mode", "multi_fill_all")
    auto_frame_fit = request.form.get("auto_frame_fit", "true") == "true"
    use_perspective = request.form.get("perspective", "true") == "true"
    crop_mode = request.form.get("crop_mode", "cover")
    boxes = json.loads(request.form.get("boxes_json", "[]"))
    deltas = json.loads(request.form.get("deltas_json", "[]"))
    if not boxes:
        boxes = [{"x": 0.2, "y": 0.2, "w": 0.6, "h": 0.6}]

    photos = request.files.getlist("photos")
    if not sid or not photos:
        return jsonify({"error": "Session ya photos missing hain"}), 400

    sdir = get_session_dir(sid)
    pptx_path = os.path.join(sdir, "template.pptx")
    prs = Presentation(pptx_path)
    sw, sh = prs.slide_width, prs.slide_height

    if fill_mode == "multi_fill_all":
        total_slides = len(prs.slides)
        p_idx = 0

        for s_idx in range(total_slides):
            if p_idx >= len(photos):
                break
            slide = prs.slides[s_idx]
            s_boxes = resolve_slide_boxes(slide, sw, sh, boxes, deltas, auto_frame_fit)
            for b in s_boxes:
                if p_idx >= len(photos):
                    break
                raw_bytes = photos[p_idx].read()
                target_ratio = (b["w"] * sw) / max(1, (b["h"] * sh))
                img_stream = process_tile_photo(raw_bytes, target_ratio, use_perspective, crop_mode)
                del raw_bytes
                place_photo_in_box(slide, sw, sh, b, img_stream)
                del img_stream
                p_idx += 1
                if p_idx % 10 == 0:
                    gc.collect()

        last_template = prs.slides[total_slides - 1]
        last_boxes = resolve_slide_boxes(last_template, sw, sh, boxes, deltas, auto_frame_fit)
        while p_idx < len(photos):
            nslide = duplicate_slide(prs, last_template)
            for b in last_boxes:
                if p_idx >= len(photos):
                    break
                raw_bytes = photos[p_idx].read()
                target_ratio = (b["w"] * sw) / max(1, (b["h"] * sh))
                img_stream = process_tile_photo(raw_bytes, target_ratio, use_perspective, crop_mode)
                del raw_bytes
                place_photo_in_box(nslide, sw, sh, b, img_stream)
                del img_stream
                p_idx += 1
                if p_idx % 10 == 0:
                    gc.collect()
    else:
        template_slide = prs.slides[template_idx]
        per_page = 1 if fill_mode == "single_repeat" else len(boxes)
        active_boxes = [boxes[0]] if fill_mode == "single_repeat" else boxes

        new_slide_elems = []
        for start in range(0, len(photos), per_page):
            chunk = photos[start:start + per_page]
            nslide = duplicate_slide(prs, template_slide)

            for bi, photo in enumerate(chunk):
                b = active_boxes[bi]
                tile_name = os.path.splitext(photo.filename)[0]
                if fill_mode == "single_repeat":
                    for shp in nslide.shapes:
                        if shp.has_text_frame:
                            for para in shp.text_frame.paragraphs:
                                for run in para.runs:
                                    if "{{NAME}}" in run.text:
                                        run.text = run.text.replace("{{NAME}}", tile_name)

                raw_bytes = photo.read()
                target_ratio = (b["w"] * sw) / max(1, (b["h"] * sh))
                img_stream = process_tile_photo(raw_bytes, target_ratio, use_perspective, crop_mode)
                del raw_bytes
                place_photo_in_box(nslide, sw, sh, b, img_stream)
                del img_stream

            if start % 10 == 0:
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