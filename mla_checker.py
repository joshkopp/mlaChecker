import re, json
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH as AL

TOL = 0.05
DATE = r"^\d{1,2} [A-Z][a-z]+\.? \d{4}$"

def ps(p, attr):
    v = getattr(p.paragraph_format, attr); s = p.style
    while v is None and s is not None:
        v = getattr(s.paragraph_format, attr); s = s.base_style
    return v

def rf(p, r, attr):
    v = getattr(r.font, attr); s = p.style
    while v is None and s is not None:
        v = getattr(s.font, attr); s = s.base_style
    return v

def inch(x): return None if x is None else x / 914400
def fancy(p): return any(r.bold or r.underline for r in p.runs) or bool(p.style.font.bold)
def new_page(P, i):
    x = P[i]._p.xml
    return bool(ps(P[i], "page_break_before")) or 'w:type="page"' in x or "lastRenderedPageBreak" in x or (i > 0 and 'w:type="page"' in P[i-1]._p.xml)

def analyze(texts, doc=None):
    P = doc.paragraphs if doc else None
    out = []
    def add(level, msg, para=None, frag=None):
        out.append(dict(level=level, msg=msg, para=para, frag=frag))
    ne = [i for i, t in enumerate(texts) if t.strip()]
    if len(ne) < 6:
        add("ERROR", "Too short to check. Is this the full essay?"); return out

    if doc:
        s = doc.sections[0]
        for n, v in [("Top", s.top_margin), ("Bottom", s.bottom_margin), ("Left", s.left_margin), ("Right", s.right_margin)]:
            if v is None or abs(inch(v) - 1) > TOL:
                add("ERROR", f'{n} margin is {"unknown" if v is None else f"{inch(v):.2f}"}". MLA needs 1" on all sides.')
        fonts, sizes = set(), set()
        for p in P:
            for r in p.runs:
                if r.text.strip():
                    fonts.add(rf(p, r, "name")); z = rf(p, r, "size"); sizes.add(z.pt if z else None)
        bad = sorted(f for f in fonts if f and f != "Times New Roman")
        if bad: add("ERROR", f"Font is {', '.join(bad)}. Use Times New Roman.")
        bs = sorted(z for z in sizes if z and z != 12)
        if bs: add("ERROR", f"Font size is {', '.join(str(z) for z in bs)} pt. Use 12 pt.")
        hx = "".join(p._p.xml for p in s.header.paragraphs)
        ht = " ".join(p.text for p in s.header.paragraphs).strip()
        if not ht and "PAGE" not in hx:
            add("ERROR", "No header. Add your last name + page number, top right (e.g. Smith 1).")
        else:
            if "PAGE" not in hx: add("ERROR", "Header has no automatic page number (Insert > Page Number).")
            if not re.search(r"[A-Za-z]{2,}", ht): add("ERROR", "Header is missing your last name before the page number.")
            if s.header.paragraphs[0].alignment != AL.RIGHT: add("WARN", "Header should be right-aligned.")

    h = ne[:4]
    if not re.match(DATE, texts[h[3]].strip()):
        add("ERROR", "4th line should be the date as Day Month Year (e.g. 3 October 2026).", h[3])
    if doc:
        for i in h:
            if P[i].alignment in (AL.CENTER, AL.RIGHT):
                add("ERROR", "Heading lines (name/instructor/course/date) should be left-aligned.", i)
    ti = ne[4]
    if doc:
        if P[ti].alignment != AL.CENTER: add("ERROR", "Title should be centered.", ti)
        if fancy(P[ti]): add("ERROR", "Title shouldn't be bold or underlined.", ti)
    if texts[ti].strip().endswith("."): add("WARN", "Title shouldn't end with a period.", ti)

    wc = next((i for i in ne if texts[i].strip().lower() in ("works cited", "work cited")), None)
    end = wc if wc is not None else len(texts)
    cited = False
    for i in ne:
        if not ti < i < end: continue
        t = texts[i]
        if doc and len(t) > 120:
            ls = ps(P[i], "line_spacing")
            if not isinstance(ls, float) or abs(ls - 2) > .05: add("ERROR", "Not double spaced.", i)
            fi = ps(P[i], "first_line_indent")
            if fi is None or abs(inch(fi) - .5) > TOL: add("ERROR", 'Missing 0.5" first-line indent.', i)
            sa = ps(P[i], "space_after")
            if sa and sa > Pt(1): add("ERROR", "Extra space after paragraph. Set it to 0 pt.", i)
        for c in re.findall(r"\(([^()]*\d[^()]*)\)", t):
            cited = True; f = f"({c})"
            if re.search(r"\b(p|pp|pg|page)\b\.?", c, re.I): add("ERROR", 'Drop "p."/"pg". Just (Author 42).', i, f)
            elif re.match(r"[A-Z][\w'\- ]*,?\s*(19|20)\d{2}", c): add("ERROR", "Looks like APA (has a year). MLA uses (Author page).", i, f)
            elif re.search(r"[A-Za-z],\s*\d", c): add("ERROR", "No comma between author and page number.", i, f)
    if not cited: add("WARN", "No in-text citations found. MLA expects (Author page), e.g. (Smith 42).")

    if wc is None:
        add("ERROR", 'No "Works Cited" page found.')
    else:
        if texts[wc].strip() != "Works Cited": add("ERROR", 'Heading should read exactly "Works Cited".', wc)
        if doc:
            if P[wc].alignment != AL.CENTER: add("ERROR", "Works Cited heading should be centered.", wc)
            if fancy(P[wc]): add("ERROR", "Works Cited heading shouldn't be bold or underlined.", wc)
            if not new_page(P, wc): add("ERROR", "Works Cited should start on its own page.", wc)
        ents = [i for i in ne if i > wc]
        if not ents: add("ERROR", "Works Cited has no entries.", wc)
        prev = ""
        for i in ents:
            t = texts[i]
            if doc:
                li, fi = ps(P[i], "left_indent"), ps(P[i], "first_line_indent")
                if not (fi is not None and fi < 0 and li is not None and abs(inch(li) - .5) <= TOL):
                    add("ERROR", 'Needs a 0.5" hanging indent.', i)
            k = re.sub(r"^(the|a|an)\s+", "", t.strip().lower())
            if k < prev: add("ERROR", "Out of alphabetical order.", i)
            prev = k
            m = re.search(r"https?://\S+", t)
            if m: add("WARN", "MLA 9 drops https://. Start at www. or the domain.", i, m.group(0))
    return out

def run_docx(path):
    doc = Document(path)
    texts = [p.text for p in doc.paragraphs]
    return json.dumps(dict(paras=texts, issues=analyze(texts, doc), formatting=True))

def run_text(s):
    texts = [l for l in s.splitlines() if l.strip()]
    return json.dumps(dict(paras=texts, issues=analyze(texts), formatting=False))
