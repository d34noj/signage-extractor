"""
Store Signage Schedule Extractor
--------------------------------
Run:   streamlit run signage_extractor.py
Needs: pip install streamlit pdfplumber pandas openpyxl pillow anthropic

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


# --------------------------------------------------------------------------------------
# UI
# --------------------------------------------------------------------------------------
def main():
    st.set_page_config(page_title="Store Signage Schedule Extractor", page_icon="📋", layout="wide")
    st.title("Store Signage Schedule Extractor")
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
        d1, d2, d3 = st.columns(3)
        d1.download_button("Download CSV", df.to_csv(index=False).encode("utf-8-sig"),
                           "signage_schedule.csv", "text/csv")
        d2.download_button("Download Excel (with snippets)", build_xlsx(df, res["snips"]),
                           "signage_schedule.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        d3.download_button("Download all snippets (ZIP)", build_zip(res["snips"]),
                           "signage_snippets.zip", "application/zip")
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


if __name__ == "__main__":
    main()
