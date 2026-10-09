"""
Store Signage Schedule Extractor
--------------------------------
Run:   streamlit run signage_extractor.py
Needs: pip install streamlit pdfplumber pandas openpyxl pillow reportlab anthropic
       (pypdfium2 comes with pdfplumber)

Two modes, switched at the top of the page:
  TK Maxx - schedule extractor (below).
  M&S     - sign register built from the Geetee sign manual: one row / one page per sign sheet, with
            qty, height, spec, flags and a position snippet cut from M&S's own elevation or plan.

How it works
  1. SCHEDULE   - reads the text of the TK_Signage schedule page (NOT table extraction,
                  which collapses these Revit sheets into one big blob) and pulls
                  Level / Description / Sign Code / Qty / MOD from each line.
  2. LOCATE     - finds every sign code on the plan and elevation sheets, with its
                  position on the page, the sheet number and the nearest view title.
  3. SNIPPETS   - crops the elevation (and plan) around each code, with the code boxed
                  in red, so you can see it in context.
  4. CHECK      - compares each sign's schedule qty with how many times it is tagged on
                  that level's signage plan, and lists anything it couldn't place.
  5. OPTIONAL   - sends the elevation crops to Claude to read size + FFL where they are
                  dimensioned (needs an API key). Anything not shown is left blank.
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
# Patterns - tweak here if a pack uses a different code format
# --------------------------------------------------------------------------------------
# TK04-22-F, TK03-02, TK56, TK05, TK04-62-IL, TKM-34IL, TK_Phone
CODE_PATTERN = r"(?:TK\d{2}(?:-\d{1,3}[A-Z]{0,2})?(?:-[A-Z]{1,2})?|TKM-\d+[A-Z]*|TK_[A-Za-z]+)"
# "L00 UK_150_Unmissable deals: Grey TK04-22-F 1 MOD6"
ROW_RE = re.compile(rf"^(L\d{{2}})\s+(.+?)\s+({CODE_PATTERN})\s+(\d+)\b(.*)$")


# --------------------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------------------
def is_schedule_page(text: str) -> bool:
    t = text.lower()
    return "tk code" in t and "family and type" in t


def classify_page(text: str):
    m = re.search(r"\b(L\d{2})[_ ]Signage Plan", text, re.I)
    if m:
        return "plan", m.group(1).upper()
    if "elevation" in text.lower():
        return "elevation", ""
    return "other", ""


def clean_description(raw: str) -> str:
    family, _, typ = raw.partition(": ")
    family, typ = family.strip(), typ.strip()
    if not typ or typ == family:
        return family
    return f"{family} - {typ}"


def size_from_name(desc: str) -> str:
    """Only what the family name itself states. Drawing sizes come from the vision step."""
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
    def dist(w):  # view titles normally sit BELOW their view, so penalise titles above the code
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
# Core analysis (cached so sliders/buttons don't re-read the PDF every click)
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

        # 1. SCHEDULE
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

        # FFL / height notes anywhere in the pack
        for i, t in texts.items():
            for line in t.splitlines():
                if re.search(r"\bFFL\b", line) and line.strip() not in notes:
                    notes.append(line.strip())

        # 2. LOCATE codes on the other sheets
        codes = {r["code"] for r in rows}
        hits = defaultdict(list)
        for i, page in enumerate(pdf.pages, 1):
            kind, plevel = classify_page(texts[i]) if i not in sched_pages else ("schedule", "")
            sheet = ""
            c = Counter(re.findall(r"\bA\d{3}\b", texts[i]))
            if c:
                sheet = c.most_common(1)[0][0]
            n_codes_here = 0
            if i not in sched_pages and any(code in texts[i] for code in codes):
                words = page.extract_words()
                for w in words:
                    tok = w["text"].strip(".,;:()[]")
                    if tok in codes:
                        n_codes_here += 1
                        h = {"page": i, "kind": kind, "level": plevel, "sheet": sheet,
                             "x0": w["x0"], "x1": w["x1"], "top": w["top"], "bottom": w["bottom"], "view": ""}
                        if kind == "elevation":
                            h["view"] = nearest_view_title(words, h)
                        hits[tok].append(h)
            page_rows.append({"page": i, "type": kind, "level": plevel, "sheet": sheet,
                              "chars": len(texts[i]), "codes_found": n_codes_here})

        # 3. SNIPPETS + 4. CHECKS
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

    return {"rows": out_rows, "snips": snips, "unmatched": unmatched, "notes": notes,
            "pages": page_rows, "schedule_pages": sched_pages, "fallback": fallback,
            "hits": [{"code": c, **{k: v for k, v in h.items() if k in ("page", "kind", "level", "sheet", "view")}}
                     for c, hl in hits.items() for h in hl]}


analyse = st.cache_data(show_spinner="Reading drawing pack...")(_analyse)


# --------------------------------------------------------------------------------------
# Optional: read size + FFL from elevation crops with Claude
# --------------------------------------------------------------------------------------
VISION_PROMPT = """This is a crop from a store signage elevation drawing. The red box marks the tag "{code}" ({desc}).
Read ONLY what is drawn or written on the drawing. Do not infer or assume.
Return JSON only, with these keys:
  "position": short phrase - which wall/area/view this sign is on
  "size": this sign's size as dimensioned (with units), or null
  "ffl": its height from finished floor level as dimensioned, and what it is measured to (e.g. "2450 to underside"), or null
  "tbc": any "TBC", "by others" or similar note attached to it, or null
  "confidence": "high", "medium" or "low"
  "evidence": the exact dimension text you used, or null
Use null for anything not shown. General drawing notes that may apply:
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


# --------------------------------------------------------------------------------------
# Output builders
# --------------------------------------------------------------------------------------
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
    """Landscape A4 print pack: optional schedule summary, then one page per sign with its snippets."""
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
# M&S (Geetee sign manual) - one sheet per sign
#   The M&S pack is not a schedule. It is a manual: one A3 sheet per sign with the title block,
#   spec callouts, "N No. required", and the position cut from M&S's elevation/plan.
#   So this tab builds a REGISTER from the sheets (one row per sheet) + a position snippet.
#   Text is read with pypdfium2 (fast on big vector drawings); nothing is guessed.
# --------------------------------------------------------------------------------------
MS_TB_RECT = (0.40, 0.0, 1.0, 0.125)  # title block, fractions of page (x0, y0 from bottom, x1, y1)

SPEC_RE = re.compile(r"acrylic|vinyl|alumin|\bLED\b|\bRAL\b|Mactac|\b3M\b|foamex|returns?\b|coil|fixings?|"
                     r"bracket|catenary|screw|sleeve|baseplate|slats?|laminated|digitally|CAD cut|stud|"
                     r"sanding|Signfix|box section|backing|carcass|plinth|gusset", re.I)
BOILER_RE = re.compile(r"property of Geetee|reproduced|permission|All measurements|REV DATE|DESCRIPTION:|"
                       r"CLIENT:|PROJECT:|ITEM REF|DRAWING NUMBER|JOB NO|DRAWN BY|Geeteesigns|^SCALE|^DATE", re.I)
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
    """Revision table text comes out scrambled, so only trust counts, letters and dates."""
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
        if BOILER_RE.search(ln) or re.fullmatch(r"[A-F]", ln) or "XXXXXX" in ln or "easuees" in ln \
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
    """Embedded pictures (the M&S elevation / plan renders) as fractions of the page, top-origin."""
    import pypdfium2.raw as pr
    W, H = page.get_size()
    boxes = []
    try:
        for o in page.get_objects(filter=[pr.FPDF_PAGEOBJ_IMAGE], max_depth=max_depth):
            l, b, r, t = o.get_bounds()
            if (r - l) * (t - b) / (W * H) < 0.02:
                continue
            if l > 0.84 * W and b < 0.14 * H:  # logo / title block
                continue
            bx = (max(0.0, l / W), max(0.0, 1 - t / H), min(1.0, r / W), min(1.0, 1 - b / H))
            if bx[2] - bx[0] < 0.03 or bx[3] - bx[1] < 0.03:  # masks / clipped junk
                continue
            boxes.append(bx)
    except Exception:
        pass
    return boxes


def ms_default_rect(boxes):
    """Default position crop = the lowest band of pictures (position views sit under the artwork)."""
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
    if "M&S Dundrum" in ls:
        info["description"] = " ".join(ls[ls.index("M&S Dundrum") + 1:]).strip()
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

    # quantities
    qty_lines, qty_values = [], []
    for l in body_lines:
        if not l.startswith("("):  # "(1 No. each floor)" only explains a total already stated
            for mm in QTY_RE.finditer(l):
                qty_values.append(int(mm.group(1)))
        if QTY_RE.search(l) and "XXXXXX" not in l and "DJ" not in l.split():
            qty_lines.append(l)
    qty_lines = list(dict.fromkeys(qty_lines))
    info["qty_lines"], info["qty_values"] = qty_lines, qty_values

    # heights / FFL
    blob = " ".join(body_lines)
    heights = []
    for mm in re.finditer(r"(\d[\d,\.]*)\s*(?:mm)?\s*(AFFL|FFL)\b([^|]{0,32})", blob, re.I):
        val = mm.group(1)
        tail = mm.group(3).strip()
        tail = tail if re.match(r"^(to|from)\b", tail, re.I) else ""
        if re.fullmatch(r"\d\.\d{3}", val):  # datum like "FFL 3.144 m"
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
    # sheet x of y per drawing number
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
    """Returns (qty or None, note). Never guesses: ambiguous sheets stay blank and get flagged."""
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


def ms_render(doc, page_no, rect, dpi, box_label=None):
    """Render a rectangle (fractions x0,y0,x1,y1 with y measured from the TOP) of a page to PNG bytes."""
    page = doc[page_no - 1]
    W, H = page.get_size()
    x0, y0, x1, y1 = rect
    x0, y0 = max(0.0, min(x0, 0.98)), max(0.0, min(y0, 0.98))
    x1, y1 = min(1.0, max(x1, x0 + 0.02)), min(1.0, max(y1, y0 + 0.02))
    im = page.render(scale=dpi / 72.0, crop=(x0 * W, (1 - y1) * H, (1 - x1) * W, y0 * H)).to_pil().convert("RGB")
    buf = io.BytesIO()
    im.save(buf, format="PNG")
    return buf.getvalue()


MS_VISION_PROMPT = """This is one A3 sheet from a sign manual (drawing {dn}, "{desc}").
Read ONLY what is drawn or written on the sheet. Do not infer or assume.
Return JSON only with these keys:
  "size": the sign's overall size as dimensioned, e.g. "801 x 275 mm", or null
  "height": its mounting height as dimensioned (AFFL/FFL or a bare height dimension on the elevation) and what it is measured to, or null
  "position": one short phrase saying where it goes (wall / area / view shown), or null
  "qty": the quantity written on the sheet, as text, or null
  "tbc": any TBC / on hold / by others note, or null
  "confidence": "high", "medium" or "low"
  "evidence": the exact dimension or note text you relied on, or null
Use null for anything not shown."""


def ms_vision_read(client, model, doc, p):
    png = ms_render(doc, p["page"], (0, 0, 1, 1), 110)
    msg = client.messages.create(
        model=model, max_tokens=600,
        messages=[{"role": "user", "content": [
            {"type": "image", "source": {"type": "base64", "media_type": "image/png",
                                         "data": base64.b64encode(png).decode()}},
            {"type": "text", "text": MS_VISION_PROMPT.format(dn=p["drawing_no"], desc=p["description"])}]}])
    text = "".join(b.text for b in msg.content if getattr(b, "type", "") == "text")
    a, b = text.find("{"), text.rfind("}")
    return json.loads(text[a:b + 1])


MS_COLS = ["Done", "Page", "Section", "Drawing No", "Description", "Kind", "Rev", "Qty (final)", "Qty as drawn",
           "Height / FFL (final)", "Dimensions seen (mm)", "Size (read)", "Illumination", "Position (captions)",
           "Spec", "Flags", "Qty note", "Notes"]


def ms_build_df(pages, vision):
    rows = []
    for p in pages:
        if p["kind"] in ("Cover", "Section divider"):
            continue
        v = vision.get(p["page"], {})
        q, qnote = ms_proposed_qty(p)
        if q is None and v.get("qty"):
            m = re.search(r"\d+", str(v["qty"]))
            if m and p["kind"] == "Sign sheet":
                q, qnote = int(m.group(0)), "qty read by Claude from sheet - check"
        height = "; ".join(p["heights"]) or (v.get("height") or "")
        flags = list(p["flags"])
        if v.get("tbc"):
            flags.append(str(v["tbc"]))
        if p["kind"] == "Sign sheet" and not p["captions"] and not p["boxes"]:
            flags.append("no position view on this sheet - see reference sheets")
        if p["kind"] == "Sign sheet" and q is None:
            flags.append("no qty on sheet")
        rows.append({
            "Done": False, "Page": p["page"], "Section": p["section"], "Drawing No": p["drawing_no"],
            "Description": p["description"] + (f"  (sheet {p['sheet_of']})" if p["sheet_of"] and not p["sheet_of"].startswith("1 of 1") else ""),
            "Kind": p["kind"], "Rev": f"{p['rev']} ({p['rev_date']})" if p["rev"] else "",
            "Qty (final)": q if p["kind"] == "Sign sheet" else None,
            "Qty as drawn": "; ".join(p["qty_lines"]),
            "Height / FFL (final)": height, "Dimensions seen (mm)": p["dims"],
            "Size (read)": v.get("size") or "", "Illumination": p["illum"],
            "Position (captions)": "; ".join(p["captions"]) or (v.get("position") or ""),
            "Spec": " / ".join(p["spec"][:6]), "Flags": "; ".join(dict.fromkeys(flags)),
            "Qty note": qnote, "Notes": "",
        })
    return pd.DataFrame(rows, columns=MS_COLS)


def ms_snip_png(doc, page_no, rects, dpi, default_rect):
    return ms_render(doc, page_no, rects.get(page_no, default_rect), dpi)


def ms_build_xlsx(df, snips) -> bytes:
    from openpyxl import Workbook
    from openpyxl.drawing.image import Image as XLImage
    from openpyxl.styles import Alignment, Font
    wb = Workbook()
    ws = wb.active
    ws.title = "M&S Sign Register"
    headers = MS_COLS + ["Position snippet"]
    ws.append(headers)
    for c in ws[1]:
        c.font = Font(bold=True)
    widths = [6, 6, 20, 13, 34, 20, 12, 9, 28, 24, 22, 18, 16, 36, 48, 30, 26, 24, 50]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = w
    keep = []
    for n, (_, row) in enumerate(df.iterrows(), start=2):
        ws.append([("" if (isinstance(row[c], float) and pd.isna(row[c])) else row[c]) for c in MS_COLS])
        for c in ws[n]:
            c.alignment = Alignment(wrap_text=True, vertical="top")
        png = snips.get(int(row["Page"]))
        if png:
            im = Image.open(io.BytesIO(png))
            im.thumbnail((380, 260))
            b = io.BytesIO()
            im.save(b, format="PNG")
            b.seek(0)
            keep.append(b)
            ws.add_image(XLImage(b), f"{ws.cell(row=1, column=len(headers)).column_letter}{n}")
            ws.row_dimensions[n].height = 200
    ws.freeze_panes = "A2"
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def ms_build_zip(df, snips) -> bytes:
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for _, r in df.iterrows():
            png = snips.get(int(r["Page"]))
            if png:
                name = re.sub(r"[^A-Za-z0-9]+", "_", r["Description"])[:40].strip("_")
                z.writestr(f"p{int(r['Page']):02d}_{r['Drawing No']}_{name}.png", png)
    return out.getvalue()


def ms_build_pdf(df, pages_by_no, snips, source_name="") -> bytes:
    """Landscape A4 check pack: summary table, then one page per sheet with facts + position snippet."""
    from xml.sax.saxutils import escape
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.platypus import (Image as RLImage, PageBreak, Paragraph, SimpleDocTemplate,
                                    Spacer, Table, TableStyle)
    ss = getSampleStyleSheet()
    h1 = ParagraphStyle("h1", parent=ss["Title"], fontSize=19, leading=22, alignment=0, spaceAfter=2)
    h2 = ParagraphStyle("h2", parent=ss["Normal"], fontSize=11, leading=14, spaceAfter=5, textColor=colors.HexColor("#444444"))
    cell = ParagraphStyle("cell", parent=ss["Normal"], fontSize=7.5, leading=9)
    cellb = ParagraphStyle("cellb", parent=cell, fontName="Helvetica-Bold")
    fact = ParagraphStyle("fact", parent=ss["Normal"], fontSize=9, leading=11.5)
    cap = ParagraphStyle("cap", parent=ss["Normal"], fontSize=8, leading=10, textColor=colors.grey)

    def P(t, s=cell):
        return Paragraph(escape(str(t if t is not None and not (isinstance(t, float) and pd.isna(t)) else "")), s)

    page_w, page_h = landscape(A4)
    margin = 28
    frame_w = page_w - 2 * margin
    img_h = page_h - 2 * margin - 150

    def footer(c, d):
        c.saveState()
        c.setFont("Helvetica", 7.5)
        c.setFillColor(colors.grey)
        c.drawString(margin, 14, f"M&S sign register - {source_name}")
        c.drawRightString(page_w - margin, 14, f"Page {d.page}")
        c.restoreState()

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(A4), leftMargin=margin, rightMargin=margin,
                            topMargin=margin, bottomMargin=margin, title="M&S sign register")
    story = [Paragraph("M&S Sign Register", h1), Paragraph(escape(source_name), h2)]
    head = ["Pg", "Drawing", "Description", "Qty", "Height / FFL", "Rev", "Flags"]
    data = [[P(h, cellb) for h in head]]
    for _, r in df.iterrows():
        q = r["Qty (final)"]
        data.append([P(r["Page"]), P(r["Drawing No"]), P(r["Description"]),
                     P("" if pd.isna(q) else int(q)), P(r["Height / FFL (final)"]), P(r["Rev"]), P(r["Flags"])])
    t = Table(data, colWidths=[24, 64, 220, 30, 150, 70, 270], repeatRows=1)
    t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#DDDDDD")),
                           ("GRID", (0, 0), (-1, -1), 0.4, colors.grey), ("VALIGN", (0, 0), (-1, -1), "TOP"),
                           ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F5F5F5")])]))
    story += [t, PageBreak()]

    def fit(png, mw, mh):
        im = Image.open(io.BytesIO(png)).convert("RGB")
        im.thumbnail((1800, 1200))  # keeps the print pack a sensible size
        jb = io.BytesIO()
        im.save(jb, format="JPEG", quality=85)
        jb.seek(0)
        w, h = im.size
        k = min(mw / w, mh / h)
        return RLImage(jb, width=w * k, height=h * k)

    n_rows = len(df)
    for n, (_, r) in enumerate(df.iterrows(), start=1):
        q = r["Qty (final)"]
        qtxt = "-" if pd.isna(q) else int(q)
        story.append(Paragraph(f"{escape(r['Description'])} &nbsp;&nbsp;<font size=11 color='#555555'>"
                               f"{escape(r['Drawing No'])} &nbsp;|&nbsp; p{int(r['Page'])} &nbsp;|&nbsp; Qty {qtxt}</font>", h1))
        story.append(Paragraph(escape(f"{r['Section']}  -  {r['Kind']}  -  Rev {r['Rev'] or '-'}"), h2))

        def F(label, value):
            return Paragraph(f"<b>{label}:</b> {escape(str(value or '-'))}", fact)
        facts = Table([[F("Height / FFL", r["Height / FFL (final)"]), F("Qty as drawn", r["Qty as drawn"])],
                       [F("Dimensions seen", r["Dimensions seen (mm)"]), F("Illumination", r["Illumination"])],
                       [F("Spec", r["Spec"]), F("Flags", r["Flags"] or r["Qty note"])]],
                      colWidths=[frame_w * 0.5, frame_w * 0.5])
        facts.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0),
                                   ("BOTTOMPADDING", (0, 0), (-1, -1), 3), ("TOPPADDING", (0, 0), (-1, -1), 0)]))
        story += [facts, Spacer(1, 6)]
        png = snips.get(int(r["Page"]))
        if png:
            story += [fit(png, frame_w, img_h - 14), Paragraph("Position snippet - M&amp;S drawing, sheet p%d" % int(r["Page"]), cap)]
        else:
            story.append(Paragraph("No snippet available.", fact))
        if n < n_rows:
            story.append(PageBreak())
    doc.build(story, onFirstPage=footer, onLaterPages=footer)
    return buf.getvalue()


# --------------------------------------------------------------------------------------
# UI
# --------------------------------------------------------------------------------------
def _ms_step(d, labels):
    try:
        i = labels.index(st.session_state.get("ms_pick"))
    except ValueError:
        i = 0
    st.session_state["ms_pick"] = labels[(i + d) % len(labels)]


def run_ms():
    st.caption("Upload the M&S internal signage manual (PDF). Each sheet becomes one register row with qty, "
               "height, spec, flags and a position snippet cut from M&S's own drawing. Nothing is guessed - "
               "blanks and flags mean the sheet doesn't say.")
    with st.sidebar:
        st.header("M&S snippet settings")
        dpi = st.slider("Snippet resolution (dpi)", 100, 250, 150, 10, key="ms_dpi")
        incl_ref = st.checkbox("Also list plan / elevation reference sheets", False, key="ms_ref")
        st.header("Read size / height with Claude (optional)")
        api_key = st.text_input("Anthropic API key", value=os.getenv("ANTHROPIC_API_KEY", ""), type="password", key="ms_key")
        model = st.text_input("Model", value="claude-sonnet-5-5", key="ms_model")

    up = st.file_uploader("Drag and drop the M&S manual PDF", type=["pdf"], key="ms_up")
    if not up:
        st.info("Waiting for a PDF.")
        return
    pdf_bytes = up.getvalue()
    digest = hashlib.md5(pdf_bytes).hexdigest()
    try:
        pages = ms_analyse(pdf_bytes)
        doc = ms_open(digest, pdf_bytes)
    except Exception as e:
        st.error(f"Couldn't read that PDF: {e}")
        return
    if not any(p["drawing_no"] for p in pages):
        st.error("No Geetee title blocks found - this doesn't look like an M&S sign manual.")
        return

    vstore = st.session_state.setdefault("ms_vision", {}).setdefault(digest, {})
    rects = st.session_state.setdefault("ms_rects", {}).setdefault(digest, {})
    srcs = st.session_state.setdefault("ms_src", {}).setdefault(digest, {})
    by_page = {p["page"]: p for p in pages}

    def snip_args(pg):
        src = srcs.get(pg, pg)
        return src, tuple(rects.get(pg, by_page[src]["rect"]))

    # ---- optional vision pass
    todo = [p for p in pages if p["kind"] == "Sign sheet" and p["page"] not in vstore]
    if st.button(f"Read size & height from sheets with Claude ({len(todo)} sheets)", disabled=not (api_key and todo)):
        try:
            import anthropic
            client = anthropic.Anthropic(api_key=api_key)
        except Exception as e:
            st.error(f"Anthropic library problem: {e}")
            client = None
        if client:
            bar = st.progress(0.0)
            for n, p in enumerate(todo, 1):
                try:
                    vstore[p["page"]] = ms_vision_read(client, model, doc, p)
                except Exception as e:
                    vstore[p["page"]] = {"confidence": "failed", "evidence": str(e)[:120]}
                bar.progress(n / len(todo))
            bar.empty()

    df_all = ms_build_df(pages, vstore)
    df = df_all if incl_ref else df_all[df_all["Kind"] == "Sign sheet"].reset_index(drop=True)

    signs = df_all[df_all["Kind"] == "Sign sheet"]
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Sheets in manual", len(df_all))
    c2.metric("Sign sheets", len(signs))
    c3.metric("With a qty on the sheet", int(signs["Qty (final)"].notna().sum()))
    c4.metric("Flagged", int((signs["Flags"] != "").sum()))

    t1, t2, t3 = st.tabs(["Register", "Sign pages", "Debug"])
    with t1:
        edit_cols = ("Done", "Qty (final)", "Height / FFL (final)", "Notes")
        edited = st.data_editor(
            df, hide_index=True, key=f"ms_editor_{digest}_{incl_ref}_{len(vstore)}",
            disabled=[c for c in MS_COLS if c not in edit_cols],
            column_config={"Done": st.column_config.CheckboxColumn("Done"),
                           "Qty (final)": st.column_config.NumberColumn("Qty (final)", min_value=0, step=1)})
        st.caption("Qty (final), Height / FFL and Notes are yours to edit - the Excel and print pack use what you "
                   "enter. Qty is only pre-filled when the sheet itself states it; floor splits are summed and "
                   "flagged. Elevation dimensions are often pictures, not text, so heights are blank unless the "
                   "sheet has AFFL / FFL as text or you run the Claude read.")

        if st.button("Build Excel / ZIP / print pack from the table above"):
            bar = st.progress(0.0, text="Cutting snippets...")
            snips = {}
            rows = list(edited.iterrows())
            for n, (_, r) in enumerate(rows, 1):
                pg = int(r["Page"])
                src, rc = snip_args(pg)
                snips[pg] = ms_render(doc, src, rc, dpi)
                bar.progress(n / len(rows), text=f"Cutting snippets... {n}/{len(rows)}")
            bar.empty()
            st.session_state["ms_out"] = {
                "digest": digest,
                "xlsx": ms_build_xlsx(edited, snips),
                "zip": ms_build_zip(edited, snips),
                "pdf": ms_build_pdf(edited, None, snips, up.name),
            }
        out = st.session_state.get("ms_out")
        if out and out["digest"] == digest:
            d1, d2, d3, d4 = st.columns(4)
            d1.download_button("Download CSV", edited.to_csv(index=False).encode("utf-8-sig"),
                               "ms_sign_register.csv", "text/csv")
            d2.download_button("Download Excel (with snippets)", out["xlsx"], "ms_sign_register.xlsx",
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            d3.download_button("Download all snippets (ZIP)", out["zip"], "ms_snippets.zip", "application/zip")
            d4.download_button("Download print pack (PDF, A4 landscape)", out["pdf"], "ms_print_pack.pdf",
                               "application/pdf")
            st.caption("Built from the table as it was when you pressed the button - rebuild after more edits.")

    with t2:
        pool = [p for p in pages if p["kind"] == "Sign sheet" or (incl_ref and p["kind"].startswith("Reference"))]
        labels = [f"p{p['page']} | {p['drawing_no']} | {p['description']}" for p in pool]
        if st.session_state.get("ms_pick") not in labels:
            st.session_state["ms_pick"] = labels[0]
        n1, n2, n3 = st.columns([1, 1, 6])
        n1.button("Previous", on_click=_ms_step, args=(-1, labels), key="ms_prev")
        n2.button("Next", on_click=_ms_step, args=(1, labels), key="ms_next")
        pick = n3.selectbox("Sheet", labels, key="ms_pick", label_visibility="collapsed")
        p = pool[labels.index(pick)]
        pg = p["page"]
        row = df_all[df_all["Page"] == pg].iloc[0]

        left, right = st.columns([2, 3])
        with left:
            st.subheader(p["description"] or "(no description)")
            st.write(f"**{p['drawing_no']}**  |  job {p['job_no']}  |  {p['section']}  |  p{pg}"
                     + (f"  |  sheet {p['sheet_of']}" if p["sheet_of"] else ""))
            q = row["Qty (final)"]
            st.write(f"**Qty:** {'-' if pd.isna(q) else int(q)}   {('(' + row['Qty note'] + ')') if row['Qty note'] else ''}")
            if p["qty_lines"]:
                st.write("**Qty as drawn:** " + "; ".join(p["qty_lines"]))
            st.write(f"**Height / FFL:** {row['Height / FFL (final)'] or '-'}")
            st.write(f"**Dimensions seen on sheet (mm):** {p['dims'] or '-'}")
            if row["Size (read)"]:
                st.write(f"**Size (read by Claude):** {row['Size (read)']}")
            st.write(f"**Illumination:** {p['illum'] or '-'}")
            st.write(f"**Rev:** {p['rev'] or '-'} {('(' + p['rev_date'] + ')') if p['rev_date'] else ''}  "
                     f"{('- ' + p['rev_notes']) if p['rev_notes'] else ''}")
            if p["spec"]:
                st.write("**Spec callouts:**")
                for l in p["spec"]:
                    st.write("- " + l)
            if row["Flags"]:
                st.warning(row["Flags"])
            if p["same_drawing"]:
                st.caption("Other sheets for this drawing: " + ", ".join(f"p{x}" for x in p["same_drawing"]))
        with right:
            same_sec = [x["page"] for x in pages if x["kind"].startswith("Reference") and x["section"] == p["section"]]
            if not p["boxes"]:
                st.info("No picture view on this sheet. Choose a reference sheet below"
                        + (f" - suggested for this section: {', '.join('p' + str(x) for x in same_sec)}." if same_sec else "."))
            with st.expander("Snippet source and crop", expanded=not p["boxes"]):
                opts = [pg] + [x for x in [y["page"] for y in pages if y["drawing_no"]] if x != pg]
                src = st.selectbox("Take the snippet from page", opts, index=opts.index(srcs.get(pg, pg)),
                                   format_func=lambda x: f"p{x} - {by_page[x]['description']}" + (" (this sheet)" if x == pg else ""),
                                   key=f"ms_src_{digest}_{pg}")
                if src == pg:
                    srcs.pop(pg, None)
                else:
                    srcs[pg] = src
                dr = by_page[src]["rect"]
                k = f"{digest}_{pg}_{src}"
                a, b = st.columns(2)
                x0 = a.slider("Left %", 0, 100, int(round(dr[0] * 100)), key=f"ms_x0_{k}")
                x1 = a.slider("Right %", 0, 100, int(round(dr[2] * 100)), key=f"ms_x1_{k}")
                y0 = b.slider("Top %", 0, 100, int(round(dr[1] * 100)), key=f"ms_y0_{k}")
                y1 = b.slider("Bottom %", 0, 100, int(round(dr[3] * 100)), key=f"ms_y1_{k}")
                if x1 - x0 >= 2 and y1 - y0 >= 2:
                    new = (x0 / 100, y0 / 100, x1 / 100, y1 / 100)
                    if tuple(round(v, 3) for v in new) != tuple(round(v, 3) for v in dr):
                        rects[pg] = new
                    else:
                        rects.pop(pg, None)
            s_src, s_rect = snip_args(pg)
            st.image(ms_render(doc, s_src, s_rect, 130), caption=f"Position snippet - p{s_src}")
            with st.expander("View whole sheet"):
                st.image(ms_render(doc, pg, (0, 0, 1, 1), 100))

    with t3:
        st.subheader("Pages")
        st.dataframe(pd.DataFrame([{
            "page": q["page"], "kind": q["kind"], "section": q["section"], "drawing": q["drawing_no"],
            "description": q["description"], "rev": q["rev"], "pictures": len(q["boxes"]),
            "qty lines": "; ".join(q["qty_lines"]), "title block text": q["tb_raw"]} for q in pages]), hide_index=True)
        st.caption("Pictures = embedded elevation / plan renders found on the sheet (used to auto-crop the snippet).")



def run_tk():
    st.caption("Upload the store fixtures / signage drawing pack (PDF). Schedule data is read from the "
               "TK_Signage schedule; position and snippets come from the plan and elevation sheets.")

    with st.sidebar:
        st.header("Snippet settings")
        elev_n = st.slider("Elevation snippets per sign", 1, 3, 1)
        incl_plan = st.checkbox("Also include a plan crop", True)
        ehw = st.slider("Elevation crop half-width (pt)", 150, 900, 350, 25)
        ehh = st.slider("Elevation crop half-height (pt)", 150, 900, 300, 25)
        phw = st.slider("Plan crop half-width (pt)", 100, 600, 250, 25)
        phh = st.slider("Plan crop half-height (pt)", 100, 600, 200, 25)
        dpi = st.slider("Snippet resolution (dpi)", 100, 250, 150, 10)
        pdf_summary = st.checkbox("Print pack: include schedule summary page(s)", True)
        st.header("Read size / FFL with Claude (optional)")
        api_key = st.text_input("Anthropic API key", value=os.getenv("ANTHROPIC_API_KEY", ""), type="password")
        model = st.text_input("Model", value="claude-sonnet-5-5")

    up = st.file_uploader("Drag and drop the store PDF pack", type=["pdf"])
    if not up:
        st.info("Waiting for a PDF.")
        return

    pdf_bytes = up.getvalue()
    digest = hashlib.md5(pdf_bytes).hexdigest()
    try:
        res = analyse(pdf_bytes, elev_n, incl_plan, ehw, ehh, phw, phh, dpi)
    except Exception as e:
        st.error(f"Couldn't read that PDF: {e}")
        return

    if not res["rows"]:
        st.error("No schedule rows found. Open the Debug tab to see what was read from each page.")

    vstore = st.session_state.setdefault("vision", {}).setdefault(digest, {})

    # ---- optional vision pass
    todo = [r for r in res["rows"]
            if f"{r['level']}|{r['code']}" not in vstore
            and any(s["kind"] == "elevation" for s in res["snips"][f"{r['level']}|{r['code']}"])]
    if st.button(f"Read size & FFL from elevation crops with Claude ({len(todo)} signs)",
                 disabled=not (api_key and todo)):
        try:
            import anthropic
            client = anthropic.Anthropic(api_key=api_key)
        except Exception as e:
            st.error(f"Anthropic library problem: {e}")
            client = None
        if client:
            bar = st.progress(0.0)
            for n, r in enumerate(todo, 1):
                key = f"{r['level']}|{r['code']}"
                png = next(s["png"] for s in res["snips"][key] if s["kind"] == "elevation")
                try:
                    vstore[key] = vision_read(client, model, png, r, res["notes"])
                except Exception as e:
                    vstore[key] = {"confidence": "failed", "evidence": str(e)[:120]}
                bar.progress(n / len(todo))
            bar.empty()

    df = build_df(res, vstore)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Schedule rows", len(df))
    c2.metric("Total qty", int(df["Qty"].sum()) if len(df) else 0)
    c3.metric("Found on drawings", sum(1 for r in res["rows"] if r["pages"]))
    c4.metric("Need checking", int((df["Check"] != "").sum()) if len(df) else 0)
    if res["fallback"]:
        st.warning("Couldn't find the TK_Signage schedule header, so every page was scanned. Check the results carefully.")

    t1, t2, t3 = st.tabs(["Schedule", "Snippets", "Debug"])
    with t1:
        st.dataframe(df, hide_index=True)
        d1, d2, d3, d4 = st.columns(4)
        d1.download_button("Download CSV", df.to_csv(index=False).encode("utf-8-sig"),
                           "signage_schedule.csv", "text/csv")
        d2.download_button("Download Excel (with snippets)", build_xlsx(df, res["snips"]),
                           "signage_schedule.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        d3.download_button("Download all snippets (ZIP)", build_zip(res["snips"]),
                           "signage_snippets.zip", "application/zip")
        d4.download_button("Download print pack (PDF, A4 landscape)",
                           build_pdf(df, res["snips"], up.name, pdf_summary),
                           "signage_print_pack.pdf", "application/pdf")
        st.caption("Size (name) is only what the family name states. Size (drawing) and FFL are blank unless "
                   "read from an elevation - nothing is guessed.")
    with t2:
        if len(df):
            opts = [f"{r['level']} | {r['code']} | {r['desc']}" for r in res["rows"]]
            pick = st.selectbox("Sign", opts)
            r = res["rows"][opts.index(pick)]
            lst = res["snips"][f"{r['level']}|{r['code']}"]
            if not lst:
                st.write("No plan or elevation reference found for this sign.")
            for s in lst:
                st.image(s["png"], caption=s["label"])
    with t3:
        st.subheader("Pages")
        st.dataframe(pd.DataFrame(res["pages"]), hide_index=True)
        st.write(f"Schedule page(s): {res['schedule_pages'] or 'none found'}")
        st.subheader("Schedule-looking lines the parser did NOT match")
        st.write(res["unmatched"] or "None")
        st.subheader("Notes mentioning FFL")
        st.write(res["notes"] or "None")
        st.subheader("Every code hit on the drawings")
        st.dataframe(pd.DataFrame(res["hits"]), hide_index=True)


def main():
    st.set_page_config(page_title="Store Signage Schedule Extractor", page_icon="📋", layout="wide")
    st.title("Store Signage Schedule Extractor")
    if hasattr(st, "segmented_control"):
        mode = st.segmented_control("Store", ["TK Maxx", "M&S"], default="TK Maxx",
                                    label_visibility="collapsed", key="store_mode") or "TK Maxx"
    else:
        mode = st.radio("Store", ["TK Maxx", "M&S"], horizontal=True, label_visibility="collapsed", key="store_mode")
    if mode == "M&S":
        run_ms()
    else:
        run_tk()


if __name__ == "__main__":
    main()
