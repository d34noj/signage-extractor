"""
Store Signage Schedule Extractor
--------------------------------
Run:   streamlit run app.py
Needs: pip install streamlit pdfplumber pandas openpyxl pillow reportlab anthropic
       (pypdfium2 comes with pdfplumber)
"""
import base64
import hashlib
import io
import json
import math
import os
import re
import zipfile
from collections import Counter, defaultdict

import pandas as pd
import pdfplumber
import streamlit as st
from PIL import Image, ImageDraw

# --------------------------------------------------------------------------------------
# STREAMLIT CONFIG (Must be the very first st.* command!)
# --------------------------------------------------------------------------------------
st.set_page_config(page_title="Store Signage Schedule Extractor", layout="wide")

# --------------------------------------------------------------------------------------
# Patterns
# --------------------------------------------------------------------------------------
CODE_PATTERN = r"(?:TK\d{2}(?:-\d{1,3}[A-Z]{0,2})?(?:-[A-Z]{1,2})?|TKM-\d+[A-Z]*|TK_[A-Za-z]+)"
ROW_RE = re.compile(rf"^(L\d{{2}})\s+(.+?)\s+({CODE_PATTERN})\s+(\d+)\b(.*)$")


# --------------------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------------------
def is_schedule_page(text: str) -> bool:
    t = text.lower()
    return "tk code" in t and "family and type" in t


def classify_page(text: str, sheet_no: str = ""):
    t = text.lower()
    m = re.search(r"\b(L\d{2})[_ ]Signage Plan", text, re.I)
    if m:
        return "plan", m.group(1).upper()
    if any(kw in t for kw in ["elevation", "elev", "internal elevation", "backwall", "shopfront", "section"]) or re.match(r"^A[45]\d{2}", sheet_no):
        return "elevation", ""
    return "other", ""


def clean_description(raw: str) -> str:
    family, _, typ = raw.partition(": ")
    family, typ = family.strip(), typ.strip()
    if not typ or typ == family:
        return family
    return f"{family} - {typ}"


def size_from_name(desc: str) -> str:
    out = []
    m = re.search(r"(\d{3,4})\s*[x×]\s*(\d{2,4})", desc)
    if m:
        out.append(f"{m.group(1)} x {m.group(2)}")
    m = re.search(r"\b(\d{2,3})mm\b", desc)
    if m:
        out.append(f"{m.group(1)}mm")
    m = re.search(r"(?:^|_)(\d{2,3})_", desc)
    if m and not out:
        out.append(f"{m.group(1)}mm letter height (name)")
    if "roundel" in desc.lower():
        m = re.search(r"(\d{3,4})\s*$", desc)
        if m:
            out.append(f"{m.group(1)} roundel (name)")
    return "; ".join(out)


def nearest_view_title(words, hit) -> str:
    cands = [w for w in words if "elevation" in w["text"].lower() or re.match(r"^L\d{2}_", w["text"])]
    if not cands:
        return ""
    cx, cy = (hit["x0"] + hit["x1"]) / 2, (hit["top"] + hit["bottom"]) / 2
    def dist(w):
        dy = w["top"] - cy
        return math.hypot((w["x0"] + w["x1"]) / 2 - cx, dy * 3 if dy < 0 else dy)
    best = min(cands, key=dist)
    line = sorted([w for w in words if abs(w["top"] - best["top"]) < 4], key=lambda w: w["x0"])
    idx = line.index(best)
    lo = hi = idx
    while lo > 0 and line[lo]["x0"] - line[lo - 1]["x1"] < 18:
        lo -= 1
    while hi < len(line) - 1 and line[hi + 1]["x0"] - line[hi]["x1"] < 18:
        hi += 1
    return " ".join(w["text"] for w in line[lo:hi + 1])


def render_crop(page, hit, half_w, half_h, dpi) -> bytes:
    x0, top, x1, bottom = page.bbox
    cx, cy = (hit["x0"] + hit["x1"]) / 2, (hit["top"] + hit["bottom"]) / 2
    box = (max(x0, cx - half_w), max(top, cy - half_h), min(x1, cx + half_w), min(bottom, cy + half_h))
    im = page.crop(box).to_image(resolution=dpi).original.convert("RGB")
    s = dpi / 72.0
    pad = 3
    ImageDraw.Draw(im).rectangle(
        [(hit["x0"] - box[0]) * s - pad, (hit["top"] - box[1]) * s - pad,
         (hit["x1"] - box[0]) * s + pad, (hit["bottom"] - box[1]) * s + pad],
        outline=(220, 0, 0), width=3)
    buf = io.BytesIO()
    im.save(buf, format="PNG")
    return buf.getvalue()


# --------------------------------------------------------------------------------------
# Core analysis
# --------------------------------------------------------------------------------------
def _analyse(pdf_bytes, elev_n, incl_plan, ehw, ehh, phw, phh, dpi):
    rows, unmatched, notes, page_rows = [], [], [], []
    texts = {}
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for i, page in enumerate(pdf.pages, 1):
            try:
                texts[i] = page.extract_text() or ""
            except Exception:
                texts[i] = ""

        sched_pages = [i for i, t in texts.items() if is_schedule_page(t)]
        fallback = not sched_pages
        scan = list(texts) if fallback else sched_pages

        for i in scan:
            for line in texts[i].splitlines():
                line = line.strip()
                m = ROW_RE.match(line)
                if m:
                    level, raw, code, qty, tail = m.groups()
                    mod = re.search(r"\bMOD\d+\b", tail)
                    rows.append({
                        "level": level, "code": code, "desc": clean_description(raw),
                        "qty": int(qty), "mod": mod.group(0) if mod else "", "sched_page": i,
                    })
                elif not fallback and re.match(r"^L\d{2}\s", line) and "TK" in line:
                    unmatched.append(f"p{i}: {line}")

        for i, t in texts.items():
            for line in t.splitlines():
                if re.search(r"\bFFL\b", line) and line.strip() not in notes:
                    notes.append(line.strip())

        codes = {r["code"] for r in rows}
        hits = defaultdict(list)
        for i, page in enumerate(pdf.pages, 1):
            sheet = ""
            c = Counter(re.findall(r"\bA\d{3}\b", texts[i]))
            if c:
                sheet = c.most_common(1)[0][0]

            kind, plevel = classify_page(texts[i], sheet) if i not in sched_pages else ("schedule", "")
            
            n_codes_here = 0
            if i not in sched_pages and any(code in texts[i] for code in codes):
                if kind == "other":
                    kind = "elevation"

                words = page.extract_words()
                for w in words:
                    tok = w["text"].strip(".,;:()[]/\\")
                    matched_code = next((c for c in codes if c in tok or tok in c), None)
                    if matched_code:
                        n_codes_here += 1
                        h = {"page": i, "kind": kind, "level": plevel, "sheet": sheet,
                             "x0": w["x0"], "x1": w["x1"], "top": w["top"], "bottom": w["bottom"], "view": ""}
                        if kind == "elevation":
                            h["view"] = nearest_view_title(words, h)
                        hits[matched_code].append(h)
            page_rows.append({"page": i, "type": kind, "level": plevel, "sheet": sheet,
                              "chars": len(texts[i]), "codes_found": n_codes_here})

        crop_cache, snips, out_rows = {}, {}, []

        def get_crop(h, hw, hh):
            key = (h["page"], round(h["x0"]), round(h["top"]), hw, hh, dpi)
            if key not in crop_cache:
                crop_cache[key] = render_crop(pdf.pages[h["page"] - 1], h, hw, hh, dpi)
            return crop_cache[key]

        for r in rows:
            code, level = r["code"], r["level"]
            elev = [h for h in hits[code] if h["kind"] == "elevation"]
            plan = [h for h in hits[code] if h["kind"] == "plan" and h["level"] == level]
            pages_found = sorted({(h["page"], h["kind"], h["sheet"] or h["level"]) for h in hits[code]})

            chosen = [(h, "elevation") for h in elev[:elev_n]]
            if incl_plan and plan:
                chosen.append((plan[0], "plan"))
            snips[f"{level}|{code}"] = [{
                "kind": k, "page": h["page"], "sheet": h["sheet"], "view": h["view"],
                "label": f"{k} p{h['page']} {h['sheet']} {h['view']}".strip(),
                "png": get_crop(h, ehw, ehh) if k == "elevation" else get_crop(h, phw, phh),
            } for h, k in chosen]

            pos = []
            if plan:
                pos.append(f"{level} plan")
            seen = set()
            for h in elev:
                label = f"{h['sheet']} {h['view']}".strip()
                if label and label not in seen:
                    seen.add(label)
                    pos.append(label)
            n_plan = len(plan)
            if n_plan == 0:
                plan_check = "not tagged on plan"
            elif n_plan == r["qty"]:
                plan_check = "OK"
            else:
                plan_check = f"plan shows {n_plan}"
            issues = []
            if not hits[code]:
                issues.append("code not found on any drawing")
            if n_plan and n_plan != r["qty"]:
                issues.append("qty differs from plan")

            out_rows.append({**r, "size_name": size_from_name(r["desc"]),
                             "position": "; ".join(pos[:4]), "plan_check": plan_check,
                             "check": "; ".join(issues),
                             "pages": ", ".join(f"p{p} {k} {s}".strip() for p, k, s in pages_found)})

    return {
        "rows": out_rows, "snips": snips, "unmatched": unmatched, "notes": notes,
        "pages": page_rows, "schedule_pages": sched_pages, "fallback": fallback,
        "hits": [{"code": c, **{k: v for k, v in h.items() if k in ("page", "kind", "level", "sheet", "view")}}
                 for c, hl in hits.items() for h in hl]
    }


analyse = st.cache_data(show_spinner="Reading drawing pack...")(_analyse)


# --------------------------------------------------------------------------------------
# Optional: Vision read
# --------------------------------------------------------------------------------------
VISION_PROMPT = """This is a crop from a store signage elevation drawing. The red box marks the tag "{code}" ({desc}).
Read ONLY what is drawn or written on the drawing. Do not infer or assume.
Return JSON only, with these keys:
  "position": short phrase - which wall/area/view this sign is on
  "size": this sign's size as dimensioned (with units), or null
  "ffl": its height from finished floor level as dimensioned, and what it is measured to, or null
  "tbc": any "TBC", "by others" or similar note attached to it, or null
  "confidence": "high", "medium" or "low"
  "evidence": the exact dimension text you used, or null
Use null for anything not shown. General drawing notes:
{notes}"""


def vision_read(client, model, png: bytes, row: dict, notes: list):
    msg = client.messages.create(
        model=model, max_tokens=700,
        messages=[{"role": "user", "content": [
            {"type": "image", "source": {"type": "base64", "media_type": "image/png",
                                       "data": base64.b64encode(png).decode()}},
            {"type": "text", "text": VISION_PROMPT.format(
                code=row["code"], desc=row["desc"], notes="\n".join(notes[:6]) or "(none)")},
        ]}])
    text = "".join(b.text for b in msg.content if getattr(b, "type", "") == "text")
    a, b = text.find("{"), text.rfind("}")
    return json.loads(text[a:b + 1])


COLS = ["Level", "Sign Code", "Description", "Size (name)", "Size (drawing)", "Position",
        "FFL", "Qty", "Plan check", "Check", "MOD", "Read confidence", "Evidence"]


def build_df(res, vision):
    hang_note = next((n for n in res["notes"] if "hanging" in n.lower()), "")
    out = []
    for r in res["rows"]:
        v = vision.get(f"{r['level']}|{r['code']}", {})
        ffl = v.get("ffl") or ""
        if not ffl and "hanging" in r["desc"].lower() and hang_note:
            ffl = f"per note: {hang_note}"
        if v.get("tbc"):
            ffl = f"{ffl} [{v['tbc']}]".strip()
        out.append({
            "Level": r["level"], "Sign Code": r["code"], "Description": r["desc"],
            "Size (name)": r["size_name"], "Size (drawing)": v.get("size") or "",
            "Position": r["position"], "FFL": ffl, "Qty": r["qty"], "Plan check": r["plan_check"],
            "Check": r["check"], "MOD": r["mod"], "Read confidence": v.get("confidence") or "",
            "Evidence": v.get("evidence") or "",
        })
    return pd.DataFrame(out, columns=COLS)


def build_xlsx(df, snips) -> bytes:
    from openpyxl import Workbook
    from openpyxl.drawing.image import Image as XLImage
    from openpyxl.styles import Alignment, Font

    wb = Workbook()
    ws = wb.active
    ws.title = "Signage Schedule"
    headers = COLS + ["Snippet"]
    ws.append(headers)
    for c in ws[1]:
        c.font = Font(bold=True)
    widths = [7, 12, 46, 20, 18, 34, 28, 6, 18, 24, 8, 12, 24, 48]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = w
    keep = []
    for n, (_, row) in enumerate(df.iterrows(), start=2):
        ws.append([row[c] for c in COLS])
        for c in ws[n]:
            c.alignment = Alignment(wrap_text=True, vertical="top")
        key = f"{row['Level']}|{row['Sign Code']}"
        if snips.get(key):
            im = Image.open(io.BytesIO(snips[key][0]["png"]))
            im.thumbnail((360, 260))
            b = io.BytesIO()
            im.save(b, format="PNG")
            b.seek(0)
            xi = XLImage(b)
            keep.append(b)
            ws.add_image(xi, f"{ws.cell(row=1, column=len(headers)).column_letter}{n}")
            ws.row_dimensions[n].height = 200
    ws.freeze_panes = "A2"
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def build_zip(snips) -> bytes:
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for key, lst in snips.items():
            level, code = key.split("|")
            for i, s in enumerate(lst, 1):
                z.writestr(f"{level}_{code}_{s['kind']}_p{s['page']}_{i}.png", s["png"])
    return out.getvalue()


def build_pdf(df, snips, source_name="", summary=True) -> bytes:
    from xml.sax.saxutils import escape
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.platypus import (Image as RLImage, PageBreak, Paragraph, SimpleDocTemplate,
                                    Spacer, Table, TableStyle)

    ss = getSampleStyleSheet()
    h1 = ParagraphStyle("h1", parent=ss["Title"], fontSize=20, leading=24, alignment=0, spaceAfter=2)
    h2 = ParagraphStyle("h2", parent=ss["Normal"], fontSize=12, leading=15, spaceAfter=6)
    cell = ParagraphStyle("cell", parent=ss["Normal"], fontSize=7.5, leading=9)
    cellb = ParagraphStyle("cellb", parent=cell, fontName="Helvetica-Bold")
    fact = ParagraphStyle("fact", parent=ss["Normal"], fontSize=9.5, leading=12)
    cap = ParagraphStyle("cap", parent=ss["Normal"], fontSize=8, leading=10, textColor=colors.grey)

    def P(text, style=cell):
        return Paragraph(escape(str(text or "")), style)

    page_w, page_h = landscape(A4)
    margin = 28
    frame_w = page_w - 2 * margin
    img_area_h = page_h - 2 * margin - 135

    def footer(canvas, doc):
        canvas.saveState()
        canvas.setFont("Helvetica", 7.5)
        canvas.setFillColor(colors.grey)
        canvas.drawString(margin, 14, f"Signage schedule - {source_name}")
        canvas.drawRightString(page_w - margin, 14, f"Page {doc.page}")
        canvas.restoreState()

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(A4), leftMargin=margin, rightMargin=margin,
                            topMargin=margin, bottomMargin=margin, title="Signage schedule")
    story = []

    if summary:
        story.append(Paragraph("Signage Schedule", h1))
        story.append(Paragraph(escape(source_name), h2))
        head = ["Lvl", "Code", "Description", "Size", "FFL", "Qty", "Position", "Check"]
        data = [[P(h, cellb) for h in head]]
        for _, r in df.iterrows():
            data.append([P(r["Level"]), P(r["Sign Code"], cellb), P(r["Description"]),
                         P(r["Size (drawing)"] or r["Size (name)"]), P(r["FFL"]), P(r["Qty"]),
                         P(r["Position"]), P(r["Check"] or r["Plan check"])])
        t = Table(data, colWidths=[28, 60, 190, 100, 112, 26, 150, 120], repeatRows=1)
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#DDDDDD")),
            ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F5F5F5")]),
        ]))
        story.append(t)
        story.append(PageBreak())

    def fit(png, max_w, max_h):
        w, h = Image.open(io.BytesIO(png)).size
        k = min(max_w / w, max_h / h)
        return RLImage(io.BytesIO(png), width=w * k, height=h * k)

    n_rows = len(df)
    for n, (_, r) in enumerate(df.iterrows(), start=1):
        story.append(Paragraph(f"{escape(r['Sign Code'])} &nbsp;&nbsp;<font size=11 color='#555555'>"
                               f"{escape(r['Level'])} &nbsp;|&nbsp; Qty {r['Qty']}</font>", h1))
        story.append(Paragraph(escape(r["Description"]), h2))
        size = r["Size (drawing)"] or r["Size (name)"] or "-"
        def F(label, value):
            return Paragraph(f"<b>{label}:</b> {escape(str(value))}", fact)
        facts = Table([[F("Size", size), F("FFL", r["FFL"] or "-")],
                       [F("Position", r["Position"] or "-"), F("Check", r["Check"] or r["Plan check"])]],
                      colWidths=[frame_w * 0.5, frame_w * 0.5])
        facts.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0),
                                   ("BOTTOMPADDING", (0, 0), (-1, -1), 3), ("TOPPADDING", (0, 0), (-1, -1), 0)]))
        story.append(facts)
        story.append(Spacer(1, 8))

        lst = snips.get(f"{r['Level']}|{r['Sign Code']}", [])[:2]
        if not lst:
            story.append(Paragraph("No plan or elevation reference found for this sign.", fact))
        else:
            each_w = frame_w / len(lst) - 6
            cells = [[[fit(s["png"], each_w, img_area_h - 14), Paragraph(escape(s["label"]), cap)]
                     for s in lst]]
            it = Table(cells, colWidths=[frame_w / len(lst)] * len(lst))
            it.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("ALIGN", (0, 0), (-1, -1), "CENTER")]))
            story.append(it)
        if n < n_rows:
            story.append(PageBreak())

    doc.build(story, onFirstPage=footer, onLaterPages=footer)
    return buf.getvalue()


# --------------------------------------------------------------------------------------
# M&S Support Helpers
# --------------------------------------------------------------------------------------
MS_TB_RECT = (0.40, 0.0, 1.0, 0.125)

SPEC_RE = re.compile(r"acrylic|vinyl|alumin|\bLED\b|\bRAL\b|Mactac|\b3M\b|foamex|returns?\b|coil|fixings?|"
                     r"bracket|catenary|screw|sleeve|baseplate|slats?|laminated|digitally|CAD cut|stud|"
                     r"sanding|Signfix|box section|backing|carcass|plinth|gusset", re.I)
BOILER_RE = re.compile(r"property of|reproduced|permission|All measurements|REV DATE|DESCRIPTION:|"
                       r"CLIENT:|PROJECT:|ITEM REF|DRAWING NUMBER|JOB NO|DRAWN BY|^SCALE|^DATE", re.I)
CAPTION_RE = re.compile(r"elevation|position|\bplan\b|floor|lounge|corridor|fashion|foodhall|food hall", re.I)
FLAG_RE = re.compile(r"\bTBC\b|ON HOLD|re-?utilised|by others|to be removed|to be re-?used", re.I)
QTY_RE = re.compile(r"(\d+)\s*No\.", re.I)
KIND_REF_RE = re.compile(r"\bplan\b|elevations?\b|\bpositions\b", re.I)


def ms_region(page, tp, x0, y0, x1, y1):
    W, H = page.get_size()
    return tp.get_text_bounded(left=x0 * W, bottom=y0 * H, right=x1 * W, top=y1 * H)


def ms_lines(s):
    return [l.strip() for l in s.replace("\r", "\n").split("\n") if l.strip()]


def ms_parse_rev(rev_text):
    real, notes, letters = [], [], []
    for ln in ms_lines(rev_text):
        m = re.match(r"^(?:([A-F])\s+)?(\d\d\.\d\d\.\d\d)\s+([A-Z]{2})\b\s*(.*)$", ln)
        if m:
            letter, date, by, tail = m.groups()
            if date != "00.00.25" and by != "XX":
                real.append(date)
                if letter:
                    letters.append(letter)
                if tail.strip():
                    notes.append(tail.strip())
            continue
        if BOILER_RE.search(ln) or re.fullmatch(r"[A-F]", ln) or "XXXXXX" in ln \
                or ln.startswith(("permission", "REV")):
            continue
        if not re.match(r"^\d\d\.\d\d\.\d\d", ln) and len(ln) > 3:
            notes.append(ln)
    letter = ""
    if real:
        by_count = chr(ord("A") + len(real) - 1)
        letter = max([by_count] + letters)
    def key(d):
        dd, mm, yy = d.split(".")
        return (yy, mm, dd)
    latest = max(real, key=key) if real else ""
    return letter, latest, notes


def ms_image_boxes(page, max_depth=5):
    import pypdfium2.raw as pr
    W, H = page.get_size()
    boxes = []
    try:
        for o in page.get_objects(filter=[pr.FPDF_PAGEOBJ_IMAGE], max_depth=max_depth):
            l, b, r, t = o.get_bounds()
            if (r - l) * (t - b) / (W * H) < 0.02:
                continue
            if l > 0.84 * W and b < 0.14 * H:
                continue
            bx = (max(0.0, l / W), max(0.0, 1 - t / H), min(1.0, r / W), min(1.0, 1 - b / H))
            if bx[2] - bx[0] < 0.03 or bx[3] - bx[1] < 0.03:
                continue
            boxes.append(bx)
    except Exception:
        pass
    return boxes


def ms_default_rect(boxes):
    if not boxes:
        return (0.0, 0.0, 1.0, 0.885)
    low = max(b[3] for b in boxes)
    band = [b for b in boxes if b[3] >= low - 0.08]
    x0 = min(b[0] for b in band) - 0.01
    y0 = min(b[1] for b in band) - 0.015
    x1 = max(b[2] for b in band) + 0.01
    y1 = max(b[3] for b in band) + 0.03
    return (max(0.0, x0), max(0.0, y0), min(1.0, x1), min(0.885, y1))


def ms_read_page(page, n):
    tp = page.get_textpage()
    full = tp.get_text_bounded().replace("\r", "\n")
    body = ms_region(page, tp, 0, MS_TB_RECT[3], 1, 1) + "\n" + ms_region(page, tp, 0, 0, MS_TB_RECT[0], MS_TB_RECT[3])
    info = {"page": n, "drawing_no": "", "job_no": "", "description": "", "date": "", "scale": "",
            "rev": "", "rev_date": "", "rev_notes": "", "tb_raw": ""}
    mnum = re.search(r"(DJ-\d{4}-\d{3})\s+(\d{4,6})", full)
    if not mnum:
        t = full.strip()
        info["kind"] = "Cover" if len(t) > 120 else "Section divider"
        info["section_title"] = " ".join(ms_lines(t))[:60] if info["kind"] == "Section divider" else ""
        info.update({"qty_lines": [], "qty_values": [], "spec": [], "captions": [], "flags": [], "heights": [],
                     "dims": "", "illum": "", "boxes": [], "rect": (0, 0, 1, 0.885)})
        return info
    info["drawing_no"], info["job_no"] = mnum.groups()

    desc_raw = ms_region(page, tp, 0.415, 0.02, 0.57, 0.095)
    ls = ms_lines(desc_raw)
    if "M&S" in " ".join(ls):
        info["description"] = " ".join(ls[1:]).strip() if len(ls) > 1 else " ".join(ls)
    else:
        info["description"] = " ".join(l for l in ls if not BOILER_RE.search(l))[:80]
    tb = ms_region(page, tp, 0.57, 0.02, 0.72, 0.095)
    m = re.search(r"(\d\d\.\d\d\.\d\d)\s+n/a", tb)
    info["date"] = m.group(1) if m else ""
    m = re.search(r"(\S+)\s*@A3", tb)
    info["scale"] = (m.group(1) + " @A3") if m else ""
    rev_text = ms_region(page, tp, 0.715, 0.01, 0.90, 0.12)
    letter, latest, notes = ms_parse_rev(rev_text)
    info["rev"], info["rev_date"], info["rev_notes"] = letter, latest, " | ".join(dict.fromkeys(notes))
    info["tb_raw"] = desc_raw.replace("\r", " ")[:200]

    body_lines = [l for l in ms_lines(body) if not BOILER_RE.search(l)]

    qty_lines, qty_values = [], []
    for l in body_lines:
        if not l.startswith("("):
            for mm in QTY_RE.finditer(l):
                qty_values.append(int(mm.group(1)))
        if QTY_RE.search(l) and "XXXXXX" not in l and "DJ" not in l.split():
            qty_lines.append(l)
    qty_lines = list(dict.fromkeys(qty_lines))
    info["qty_lines"], info["qty_values"] = qty_lines, qty_values

    blob = " ".join(body_lines)
    heights = []
    for mm in re.finditer(r"(\d[\d,\.]*)\s*(?:mm)?\s*(AFFL|FFL)\b([^|]{0,32})", blob, re.I):
        val = mm.group(1)
        tail = mm.group(3).strip()
        tail = tail if re.match(r"^(to|from)\b", tail, re.I) else ""
        if re.fullmatch(r"\d\.\d{3}", val):
            continue
        heights.append(f"{val} {mm.group(2).upper()} {tail}".strip())
    for mm in re.finditer(r"\b(AFFL|FFL)\s+(\d{3,4})\b", blob):
        heights.append(f"{mm.group(2)} {mm.group(1)}")
    info["heights"] = list(dict.fromkeys(heights))[:4]

    nums = []
    for l in body_lines:
        if re.fullmatch(r"\d{2,4}", l) and int(l) >= 50 and l not in nums:
            nums.append(l)
    info["dims"] = ", ".join(nums[:12])

    spec = []
    for l in body_lines:
        if SPEC_RE.search(l) and len(l) > 8 and not re.match(r"^Scale", l) and "No." not in l:
            spec.append(l)
    info["spec"] = list(dict.fromkeys(spec))[:10]
    info["illum"] = "Illuminated (LED)" if re.search(r"\bLED\b", blob) else ("Non-illuminated" if info["spec"] else "")
    info["captions"] = list(dict.fromkeys(l for l in body_lines if CAPTION_RE.search(l) and len(l) < 70))[:6]
    info["flags"] = list(dict.fromkeys(l for l in body_lines if FLAG_RE.search(l)))[:5]

    desc = info["description"]
    info["kind"] = "Reference (plan / elevation)" if KIND_REF_RE.search(desc) and not re.search(r"vinyls?\b", desc, re.I) \
        else "Sign sheet"
    info["boxes"] = ms_image_boxes(page)
    info["rect"] = ms_default_rect(info["boxes"])
    return info


def _ms_analyse(pdf_bytes):
    import pypdfium2 as pdfium
    pdf = pdfium.PdfDocument(pdf_bytes)
    pages, section = [], ""
    for i in range(len(pdf)):
        info = ms_read_page(pdf[i], i + 1)
        if info["kind"] == "Section divider":
            section = info["section_title"]
        elif info["kind"] == "Cover":
            section = ""
        info["section"] = section
        pages.append(info)
    by_no = defaultdict(list)
    for p in pages:
        if p["drawing_no"]:
            by_no[p["drawing_no"]].append(p["page"])
    for p in pages:
        lst = by_no.get(p["drawing_no"], [])
        p["sheet_of"] = f"{lst.index(p['page']) + 1} of {len(lst)}" if lst else ""
        p["same_drawing"] = [x for x in lst if x != p["page"]]
    ref_pages = [p["page"] for p in pages if p["kind"].startswith("Reference")]
    for p in pages:
        p["ref_pages"] = ref_pages
    return pages


def ms_proposed_qty(p):
    vals, lines = p["qty_values"], p["qty_lines"]
    if not vals:
        return None, "no qty on sheet - count from plan/elevation"
    floor = [l for l in lines if re.search(r"ground floor|first floor", l, re.I)]
    if floor:
        fv = [int(QTY_RE.search(l).group(1)) for l in floor if QTY_RE.search(l)]
        return sum(fv), f"summed from floor split ({' + '.join(map(str, fv))})"
    if len(set(vals)) == 1:
        note = ""
        if len(vals) > 1 or any(re.search(r"sets?|each|req.d|D/S", l, re.I) for l in lines if not l.startswith("(")):
            note = "qty is per set / per item - read the sheet"
        return vals[0], note
    return None, "several different qty callouts - read the sheet"


@st.cache_data(show_spinner="Reading manual...")
def ms_analyse(pdf_bytes):
    return _ms_analyse(pdf_bytes)


@st.cache_resource(show_spinner=False)
def ms_open(digest, pdf_bytes):
    import pypdfium2 as pdfium
    return pdfium.PdfDocument(pdf_bytes)


MS_VISION_PROMPT = """This is one A3 sheet from an M&S sign manual (drawing {dn}, "{desc}").
Read ONLY what is drawn or written on the sheet. Do not infer or assume.
Return JSON only with these keys:
  "size": the sign's overall size as dimensioned, or null
  "height": its mounting height as dimensioned, or null
  "position": one short phrase saying where it goes, or null
  "qty": the quantity written on the sheet, or null
  "tbc": any TBC / on hold note, or null
  "confidence": "high", "medium" or "low"
  "evidence": exact dimension text, or null
Use null for anything not shown."""


def ms_vision_read(client, model, doc, p):
    page = doc[p["page"] - 1]
    W, H = page.get_size()
    im = page.render(scale=110 / 72.0).to_pil().convert("RGB")
    buf = io.BytesIO()
    im.save(buf, format="PNG")
    png = buf.getvalue()
    msg = client.messages.create(
        model=model, max_tokens=700,
        messages=[{"role": "user", "content": [
            {"type": "image", "source": {"type": "base64", "media_type": "image/png",
                                       "data": base64.b64encode(png).decode()}},
            {"type": "text", "text": MS_VISION_PROMPT.format(dn=p["drawing_no"], desc=p["description"])},
        ]}])
    text = "".join(b.text for b in msg.content if getattr(b, "type", "") == "text")
    a, b_idx = text.find("{"), text.rfind("}")
    return json.loads(text[a:b_idx + 1])


# --------------------------------------------------------------------------------------
# Streamlit UI Layout
# --------------------------------------------------------------------------------------
st.title("Store Signage Schedule Extractor")
st.markdown("Upload your store sign fixtures pack PDF below to parse codes, descriptions, sizes, FFLs, and generate clean snippets.")

mode = st.radio("Select Pack Type:", ["TK Maxx", "M&S"])
uploaded_file = st.file_uploader("Upload PDF Pack", type=["pdf"])

anthropic_key = st.text_input("Anthropic API Key (Optional - for AI Vision extraction)", type="password")

if uploaded_file:
    pdf_bytes = uploaded_file.read()
    digest = hashlib.md5(pdf_bytes).hexdigest()

    if mode == "TK Maxx":
        st.subheader("TK Maxx Extractor Settings")
        col1, col2, col3 = st.columns(3)
        with col1:
            elev_n = st.slider("Elevations to crop per sign", 1, 3, 1)
            incl_plan = st.checkbox("Include plan snippet", value=True)
        with col2:
            ehw = st.slider("Elevation crop width (pts)", 200, 800, 450)
            ehh = st.slider("Elevation crop height (pts)", 150, 600, 300)
        with col3:
            phw = st.slider("Plan crop width (pts)", 200, 800, 350)
            phh = st.slider("Plan crop height (pts)", 150, 600, 250)
            dpi = st.slider("Render DPI", 72, 200, 110)

        with st.spinner("Analyzing TK Maxx PDF pack..."):
            res = analyse(pdf_bytes, elev_n, incl_plan, ehw, ehh, phw, phh, dpi)

        vision_data = {}
        if anthropic_key and st.button("Run Claude Vision Extraction on Crops"):
            import anthropic
            client = anthropic.Anthropic(api_key=anthropic_key)
            prog = st.progress(0)
            items = list(res["snips"].items())
            for idx, (key, lst) in enumerate(items):
                level, code = key.split("|")
                row = next((r for r in res["rows"] if r["code"] == code and r["level"] == level), None)
                if row and lst:
                    try:
                        vision_data[key] = vision_read(client, "claude-3-5-sonnet-20241022", lst[0]["png"], row, res["notes"])
                    except Exception as e:
                        st.warning(f"Vision failed for {code}: {e}")
                prog.progress((idx + 1) / len(items))

        df = build_df(res, vision_data)
        st.success(f"Parsed {len(df)} signage items successfully!")
        st.dataframe(df, use_container_width=True)

        st.subheader("Download Reports")
        d1, d2, d3 = st.columns(3)
        with d1:
            st.download_button("Download Excel (.xlsx)", data=build_xlsx(df, res["snips"]), file_name="signage_schedule.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with d2:
            st.download_button("Download Snippets (.zip)", data=build_zip(res["snips"]), file_name="signage_snippets.zip", mime="application/zip")
        with d3:
            st.download_button("Download Summary PDF", data=build_pdf(df, res["snips"], uploaded_file.name), file_name="signage_pack.pdf", mime="application/pdf")

    else:
        st.subheader("M&S Sign Manual Register")
        pages = ms_analyse(pdf_bytes)
        doc = ms_open(digest, pdf_bytes)

        ms_rows = []
        for p in pages:
            if p["kind"] == "Sign sheet":
                q, q_note = ms_proposed_qty(p)
                ms_rows.append({
                    "Page": p["page"],
                    "Drawing No": p["drawing_no"],
                    "Description": p["description"],
                    "Qty": q if q is not None else "Check",
                    "Qty Note": q_note,
                    "Scale": p["scale"],
                    "Rev": p["rev"],
                })
        ms_df = pd.DataFrame(ms_rows)
        st.dataframe(ms_df, use_container_width=True)
        st.info("Switch to TK Maxx mode for full layout snippet cropping, or use the M&S manual register breakdown above.")
