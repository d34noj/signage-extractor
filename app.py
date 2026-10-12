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
    m = re.search(r"\b(L\d{2})[ _-]*Signage[ _-]*Plan\b", text, re.I)
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

        codes = {str(r["code"]).upper() for r in rows if r.get("code")}
        hits = defaultdict(list)
        # Match every non-schedule page. Do not pre-filter with page.extract_text():
        # CAD exports often omit or fragment code strings in the plain-text layer.
        for i, page in enumerate(pdf.pages, 1):
            sheet = ""
            c = Counter(re.findall(r"\\bA\\d{3}\\b", texts[i], re.I))
            if c:
                sheet = c.most_common(1)[0][0].upper()

            kind, plevel = classify_page(texts[i], sheet) if i not in sched_pages else ("schedule", "")
            n_codes_here = 0
            if i not in sched_pages and codes:
                try:
                    words = page.extract_words(
                        x_tolerance=2, y_tolerance=3, keep_blank_chars=False,
                        use_text_flow=False
                    )
                except Exception:
                    words = []

                # Cluster words by baseline. A code may be represented as TK21, -, 12.
                line_groups = []
                for w in sorted(words, key=lambda x: (x["top"], x["x0"])):
                    target = next((g for g in reversed(line_groups[-8:])
                                   if abs(g["top"] - w["top"]) <= 4), None)
                    if target is None:
                        target = {"top": w["top"], "words": []}
                        line_groups.append(target)
                    target["words"].append(w)

                found_on_page = []
                for g in line_groups:
                    lw = sorted(g["words"], key=lambda x: x["x0"])
                    # Avoid joining unrelated labels far apart on a drawing line.
                    clusters = []
                    for w in lw:
                        if not clusters or w["x0"] - clusters[-1][-1]["x1"] > 36:
                            clusters.append([w])
                        else:
                            clusters[-1].append(w)
                    for cluster in clusters:
                        compact_chars, char_words = [], []
                        for w in cluster:
                            token = re.sub(r"\\s+", "", str(w.get("text", "")).upper())
                            compact_chars.extend(token)
                            char_words.extend([w] * len(token))
                        compact = "".join(compact_chars)
                        if not compact:
                            continue
                        for matched_code in sorted(codes, key=len, reverse=True):
                            mc = re.sub(r"\\s+", "", matched_code.upper())
                            # Ignore matches embedded in a longer alphanumeric code.
                            pat = re.compile(r"(?<![A-Z0-9])" + re.escape(mc) + r"(?![A-Z0-9])")
                            for match in pat.finditer(compact):
                                matched_words = char_words[match.start():match.end()]
                                if not matched_words:
                                    continue
                                x0 = min(w["x0"] for w in matched_words)
                                x1 = max(w["x1"] for w in matched_words)
                                top = min(w["top"] for w in matched_words)
                                bottom = max(w["bottom"] for w in matched_words)
                                found_on_page.append((matched_code, x0, x1, top, bottom))

                # Character-level fallback catches PDFs whose word segmentation is unusable.
                if not found_on_page:
                    try:
                        chars = [ch for ch in page.chars if str(ch.get("text", "")).strip()]
                    except Exception:
                        chars = []
                    char_lines = []
                    for ch in sorted(chars, key=lambda x: (x["top"], x["x0"])):
                        target = next((g for g in reversed(char_lines[-10:])
                                       if abs(g["top"] - ch["top"]) <= 3), None)
                        if target is None:
                            target = {"top": ch["top"], "chars": []}
                            char_lines.append(target)
                        target["chars"].append(ch)
                    for g in char_lines:
                        line_chars = sorted(g["chars"], key=lambda x: x["x0"])
                        compact_chars, char_boxes = [], []
                        prev = None
                        for ch in line_chars:
                            # Large horizontal gaps separate drawing labels; keep them separate.
                            if prev is not None and ch["x0"] - prev["x1"] > 36:
                                compact_chars.append(" ")
                                char_boxes.append(None)
                            token = re.sub(r"\\s+", "", str(ch.get("text", "")).upper())
                            compact_chars.extend(token)
                            char_boxes.extend([ch] * len(token))
                            prev = ch
                        compact = "".join(compact_chars)
                        for matched_code in sorted(codes, key=len, reverse=True):
                            mc = re.sub(r"\\s+", "", matched_code.upper())
                            pat = re.compile(r"(?<![A-Z0-9])" + re.escape(mc) + r"(?![A-Z0-9])")
                            for match in pat.finditer(compact):
                                boxes = [b for b in char_boxes[match.start():match.end()] if b]
                                if boxes:
                                    found_on_page.append((
                                        matched_code,
                                        min(b["x0"] for b in boxes), max(b["x1"] for b in boxes),
                                        min(b["top"] for b in boxes), max(b["bottom"] for b in boxes)
                                    ))

                seen_hits = set()
                for matched_code, x0, x1, top, bottom in found_on_page:
                    dedupe = (matched_code, round(x0), round(top))
                    if dedupe in seen_hits:
                        continue
                    seen_hits.add(dedupe)
                    n_codes_here += 1
                    # Plan detection uses the page title and level. Other code-bearing sheets
                    # are retained as elevation candidates rather than discarded.
                    hit_kind = kind
                    if hit_kind == "other":
                        hit_kind = "elevation"
                    h = {"page": i, "kind": hit_kind, "level": plevel, "sheet": sheet,
                         "x0": x0, "x1": x1, "top": top, "bottom": bottom, "view": ""}
                    if hit_kind == "elevation":
                        h["view"] = nearest_view_title(words, h)
                    hits[matched_code].append(h)

                    # If a code is found, retain the page as a candidate in diagnostics.
                    if kind == "other":
                        kind = "elevation"

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

            # Prefer genuine elevation sheets with a nearby view title over incidental
            # references elsewhere in the pack. Stable page order breaks ties.
            elev = sorted(elev, key=lambda h: (0 if h.get("view") else 1,
                                                0 if h.get("sheet", "").startswith(("A4", "A5")) else 1,
                                                h["page"]))
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
# UI
# --------------------------------------------------------------------------------------
def get_api_key() -> str:
    try:
        if "ANTHROPIC_API_KEY" in st.secrets:
            return st.secrets["ANTHROPIC_API_KEY"]
    except Exception:
        pass
    return os.environ.get("ANTHROPIC_API_KEY", "")


def run_vision(res, keys, api_key, model, state):
    import anthropic
    client = anthropic.Anthropic(api_key=api_key)
    by_key = {f"{r['level']}|{r['code']}": r for r in res["rows"]}
    if not keys:
        st.info("There are no signs with elevation snippets to read.")
        return
    bar = st.progress(0.0, text="Reading drawings...")
    for n, key in enumerate(keys, 1):
        snips = [s for s in res["snips"].get(key, []) if s["kind"] == "elevation"]
        if snips and key not in state:
            try:
                state[key] = vision_read(client, model, snips[0]["png"], by_key[key], res["notes"])
            except Exception as e:
                state[key] = {"evidence": f"read failed: {e}"}
        bar.progress(n / len(keys), text=f"Reading drawings... {n}/{len(keys)}")
    bar.empty()


def tk_mode():
    with st.sidebar:
        st.header("Settings")
        elev_n = st.slider("Elevation snippets per sign", 1, 3, 1)
        incl_plan = st.checkbox("Also include a plan snippet", value=True)
        ehw = st.slider("Elevation crop half-width (pt)", 50, 600, 220, 10)
        ehh = st.slider("Elevation crop half-height (pt)", 50, 600, 160, 10)
        phw = st.slider("Plan crop half-width (pt)", 50, 600, 150, 10)
        phh = st.slider("Plan crop half-height (pt)", 50, 600, 110, 10)
        dpi = st.select_slider("Snippet resolution (dpi)", options=[100, 150, 200, 250], value=150)
        st.divider()
        st.subheader("Vision read (optional)")
        api_key = get_api_key() or st.text_input("Anthropic API key", type="password")
        model = st.text_input(
            "Anthropic model ID",
            value=os.environ.get("ANTHROPIC_MODEL", "claude-haiku-4-5-20251001"),
            help="Use a model ID enabled for your Anthropic account; can also be set with ANTHROPIC_MODEL."
        )

    up = st.file_uploader("Upload the TK Maxx drawing pack (PDF)", type=["pdf"], key="tk_pdf")
    if not up:
        st.info("Upload a pack to get the schedule, one page per sign, and a check sheet.")
        return

    pdf_bytes = up.getvalue()
    file_id = hashlib.md5(pdf_bytes).hexdigest()
    try:
        res = analyse(pdf_bytes, elev_n, incl_plan, ehw, ehh, phw, phh, dpi)
    except Exception as e:
        st.error(f"Couldn't read that PDF: {e}")
        return

    if not res["rows"]:
        st.error("No schedule rows found. Check this is the TK drawing pack with the 'TK Code / Family and Type' schedule.")
        return
    if res["fallback"]:
        st.warning("No schedule page recognised - scanned every page for rows instead. Check results carefully.")

    vstore = st.session_state.setdefault("vision", {})
    vision = vstore.setdefault(file_id, {})
    df = build_df(res, vision)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Signs", len(df))
    c2.metric("Total qty", int(df["Qty"].sum()))
    c3.metric("With snippets", sum(1 for k, v in res["snips"].items() if v))
    c4.metric("Flagged", int((df["Check"] != "").sum()))

    t_sched, t_signs, t_check = st.tabs(["Schedule", "Sign by sign", "Check sheet"])

    # ---- Schedule
    with t_sched:
        st.dataframe(df, use_container_width=True, hide_index=True)
        d1, d2, d3, d4 = st.columns(4)
        d1.download_button("Excel (with snippets)", build_xlsx(df, res["snips"]),
                           f"signage_schedule.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        d2.download_button("PDF (one page per sign)", build_pdf(df, res["snips"], up.name),
                           "signage_schedule.pdf", "application/pdf")
        d3.download_button("Snippets (zip)", build_zip(res["snips"]), "snippets.zip", "application/zip")
        d4.download_button("CSV", df.to_csv(index=False).encode("utf-8"), "signage_schedule.csv", "text/csv")

        st.divider()
        if api_key:
            if st.button("Read sizes / FFL from drawings (all signs)"):
                run_vision(res, list(res["snips"].keys()), api_key, model, vision)
                st.rerun()
            st.caption("Vision reads only what's drawn on the crop - always check against the evidence column.")
        else:
            st.caption("Add an Anthropic API key in the sidebar to read sizes and FFL off the drawings.")

    # ---- Sign by sign
    with t_signs:
        labels = [f"{r['Level']}  {r['Sign Code']}  -  {r['Description']}" for _, r in df.iterrows()]
        st.session_state.setdefault("sign_idx", 0)
        st.session_state["sign_idx"] = min(st.session_state["sign_idx"], len(labels) - 1)

        def step(d):
            st.session_state["sign_idx"] = (st.session_state["sign_idx"] + d) % len(labels)

        nav1, nav2, nav3 = st.columns([1, 6, 1])
        nav1.button("< Prev", on_click=step, args=(-1,), use_container_width=True)
        nav3.button("Next >", on_click=step, args=(1,), use_container_width=True)
        nav2.selectbox("Sign", range(len(labels)), format_func=lambda i: labels[i],
                       key="sign_idx", label_visibility="collapsed")

        i = st.session_state["sign_idx"]
        r = df.iloc[i]
        key = f"{r['Level']}|{r['Sign Code']}"
        st.subheader(f"{r['Sign Code']}  |  {r['Level']}  |  Qty {r['Qty']}")
        st.write(r["Description"])
        f1, f2 = st.columns(2)
        f1.markdown(f"**Size:** {r['Size (drawing)'] or r['Size (name)'] or '-'}")
        f1.markdown(f"**FFL:** {r['FFL'] or '-'}")
        f2.markdown(f"**Position:** {r['Position'] or '-'}")
        f2.markdown(f"**Check:** {r['Check'] or r['Plan check']}")
        if r["Evidence"]:
            st.caption(f"Evidence: {r['Evidence']} ({r['Read confidence']} confidence)")

        if api_key and st.button("Read this sign from the drawing", key=f"vr_{key}"):
            run_vision(res, [key], api_key, model, vision)
            st.rerun()

        snips = res["snips"].get(key, [])
        if not snips:
            st.warning("No plan or elevation reference found for this sign.")
        else:
            cols = st.columns(len(snips))
            for col, s in zip(cols, snips):
                col.image(s["png"], caption=s["label"], use_container_width=True)

    # ---- Check sheet
    with t_check:
        flagged = df[df["Check"] != ""]
        if flagged.empty:
            st.success("Every sign was found on the drawings and plan quantities agree.")
        else:
            st.markdown("**Signs to look at**")
            st.dataframe(flagged[["Level", "Sign Code", "Description", "Qty", "Plan check", "Check"]],
                         use_container_width=True, hide_index=True)
        if res["unmatched"]:
            st.markdown("**Schedule lines that looked like signs but didn't parse**")
            st.code("\n".join(res["unmatched"]))
        if res["notes"]:
            st.markdown("**FFL notes found on the drawings**")
            for n in res["notes"][:15]:
                st.write(f"- {n}")
        with st.expander("How each page was classified"):
            st.dataframe(pd.DataFrame(res["pages"]), use_container_width=True, hide_index=True)


def ms_mode():
    st.info("M&S mode isn't built into this version yet. Gemini's rebuild only contains the TK Maxx extractor. "
            "To do M&S the same way (sign per page, snippets, check sheet) I need a few pages of the M&S "
            "binder - the schedule/legend page and one or two elevations - to see how their sign codes are laid out.")


st.title("Store Signage Schedule Extractor")
mode = st.radio("Retailer", ["TK Maxx", "M&S"], horizontal=True, label_visibility="collapsed")
if mode == "TK Maxx":
    tk_mode()
else:
    ms_mode()
