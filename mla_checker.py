import re
import sys
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

TOL = 0.05  # inches of wiggle room


# ---------- helpers: look up a setting on the paragraph, then its style chain ----------
def para_setting(p, attr):
    val = getattr(p.paragraph_format, attr)
    if val is not None:
        return val
    style = p.style
    while style is not None:
        val = getattr(style.paragraph_format, attr)
        if val is not None:
            return val
        style = style.base_style
    return None


def run_font(p, run, attr):
    val = getattr(run.font, attr)
    if val is not None:
        return val
    style = p.style
    while style is not None:
        val = getattr(style.font, attr)
        if val is not None:
            return val
        style = style.base_style
    return None


def has_page_break_before(paragraphs, i):
    p = paragraphs[i]
    if para_setting(p, "page_break_before"):
        return True
    if i > 0:
        xml = paragraphs[i - 1]._p.xml
        if 'w:type="page"' in xml or "lastRenderedPageBreak" in xml:
            return True
    return 'w:type="page"' in p._p.xml


def inches(x):
    return None if x is None else x / 914400


# ---------- the checks ----------
def check_mla(path):
    doc = Document(path)
    issues = []

    def add(level, msg):
        issues.append((level, msg))

    paras = [p for p in doc.paragraphs]
    nonempty = [(i, p) for i, p in enumerate(paras) if p.text.strip()]
    if len(nonempty) < 6:
        add("ERROR", "Doc is too short to check. Is this the full essay?")
        return issues

    # 1. Margins
    s = doc.sections[0]
    for name, val in [("top", s.top_margin), ("bottom", s.bottom_margin),
                      ("left", s.left_margin), ("right", s.right_margin)]:
        if val is None or abs(inches(val) - 1.0) > TOL:
            got = "unknown" if val is None else f'{inches(val):.2f}"'
            add("ERROR", f'{name.title()} margin is {got}. MLA needs 1" on all sides.')

    # 2. Font + size (checked across all body runs)
    fonts, sizes = set(), set()
    for p in paras:
        for r in p.runs:
            if r.text.strip():
                fonts.add(run_font(p, r, "name"))
                sz = run_font(p, r, "size")
                sizes.add(sz.pt if sz else None)
    bad_fonts = {f for f in fonts if f and f != "Times New Roman"}
    if bad_fonts:
        add("ERROR", f"Font is {', '.join(sorted(bad_fonts))}. Use Times New Roman (or another readable serif).")
    bad_sizes = {z for z in sizes if z and z != 12}
    if bad_sizes:
        add("ERROR", f"Font size is {', '.join(str(z) for z in sorted(bad_sizes))} pt. Use 12 pt.")

    # 3. Header: last name + page number
    header_xml = "".join(p._p.xml for p in s.header.paragraphs)
    header_text = " ".join(p.text for p in s.header.paragraphs).strip()
    if not header_text and "PAGE" not in header_xml:
        add("ERROR", "No header found. Add your last name and page number, top right (e.g. Smith 1).")
    else:
        if "PAGE" not in header_xml:
            add("ERROR", "Header has no automatic page number. Insert one with Insert > Page Number.")
        if not re.search(r"[A-Za-z]{2,}", header_text):
            add("ERROR", "Header is missing your last name before the page number.")
        if s.header.paragraphs and s.header.paragraphs[0].alignment != WD_ALIGN_PARAGRAPH.RIGHT:
            add("WARN", "Header should be right-aligned.")

    # 4. Heading block (name, instructor, course, date) then title
    first4 = [p.text.strip() for _, p in nonempty[:4]]
    date_re = r"^\d{1,2} [A-Z][a-z]+\.? \d{4}$"
    if not re.match(date_re, first4[3]):
        add("ERROR", f'4th line "{first4[3]}" should be the date as Day Month Year (e.g. 3 October 2026).')
    for k, label in enumerate(["student name", "instructor name", "course"]):
        if len(first4[k]) < 2:
            add("ERROR", f"Line {k+1} of the heading should be your {label}.")
    for _, p in nonempty[:4]:
        if p.alignment in (WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.RIGHT):
            add("ERROR", "Heading lines (name/instructor/course/date) should be left-aligned.")
            break

    ti, title = nonempty[4]
    if title.alignment != WD_ALIGN_PARAGRAPH.CENTER:
        add("ERROR", "Title should be centered.")
    if any(r.bold or r.underline for r in title.runs) or (title.style.font.bold):
        add("ERROR", "Title should not be bold or underlined. Plain text only.")
    if title.text.strip().endswith("."):
        add("WARN", "Title shouldn't end with a period.")

    # 5. Body paragraphs: double spacing, 0.5" indent, no extra space between
    wc_idx = next((i for i, p in nonempty if p.text.strip().lower() in ("works cited", "work cited")), None)
    body_end = wc_idx if wc_idx is not None else len(paras)
    body = [p for i, p in nonempty if ti < i < body_end and len(p.text) > 120]

    bad_spacing = bad_indent = bad_after = 0
    for p in body:
        ls = para_setting(p, "line_spacing")
        if not isinstance(ls, float) and ls is not None:  # exact/at-least spacing
            bad_spacing += 1
        elif ls is None or abs(ls - 2.0) > 0.05:
            bad_spacing += 1
        fi = para_setting(p, "first_line_indent")
        if fi is None or abs(inches(fi) - 0.5) > TOL:
            bad_indent += 1
        sa = para_setting(p, "space_after")
        if sa is not None and sa > Pt(1):
            bad_after += 1
    if body:
        if bad_spacing:
            add("ERROR", f"{bad_spacing} of {len(body)} body paragraphs aren't double spaced.")
        if bad_indent:
            add("ERROR", f'{bad_indent} of {len(body)} body paragraphs lack a 0.5" first-line indent.')
        if bad_after:
            add("ERROR", f"{bad_after} body paragraphs have extra space after them. Set it to 0 pt.")
    else:
        add("WARN", "Couldn't find any body paragraphs to check spacing/indents.")

    # 6. In-text citations
    body_text = " ".join(p.text for i, p in nonempty if ti < i < body_end)
    parens = re.findall(r"\(([^()]*\d[^()]*)\)", body_text)
    if not parens:
        add("WARN", "No in-text citations found. MLA expects (Author page), e.g. (Smith 42).")
    for c in parens:
        if re.search(r"\b(p|pp|pg|page)\b\.?", c, re.I):
            add("ERROR", f'Citation ({c}): drop "p."/"pg"/"page". Just (Author 42).')
        if re.search(r",\s*\d", c) and not re.search(r"\d{4}", c):
            add("ERROR", f"Citation ({c}): no comma between author and page number.")
        if re.search(r"\b(19|20)\d{2}\b", c) and re.search(r"[A-Za-z]", c):
            add("ERROR", f"Citation ({c}): looks like APA (has a year). MLA uses author + page.")

    # 7. Works Cited
    if wc_idx is None:
        add("ERROR", 'No "Works Cited" page found.')
    else:
        wc = paras[wc_idx]
        if wc.text.strip() != "Works Cited":
            add("ERROR", 'Heading should read exactly "Works Cited".')
        if wc.alignment != WD_ALIGN_PARAGRAPH.CENTER:
            add("ERROR", "Works Cited heading should be centered.")
        if any(r.bold or r.underline for r in wc.runs):
            add("ERROR", "Works Cited heading should not be bold or underlined.")
        if not has_page_break_before(paras, wc_idx):
            add("ERROR", "Works Cited should start on its own page.")

        entries = [p for p in paras[wc_idx + 1:] if p.text.strip()]
        if not entries:
            add("ERROR", "Works Cited page has no entries.")
        for p in entries:
            li = para_setting(p, "left_indent")
            fi = para_setting(p, "first_line_indent")
            hanging = fi is not None and fi < 0 and li is not None and abs(inches(li) - 0.5) <= TOL
            if not hanging:
                add("ERROR", f'Entry needs a 0.5" hanging indent: "{p.text[:40]}..."')
                break
        keys = [re.sub(r"^(the|a|an)\s+", "", p.text.strip().lower()) for p in entries]
        if keys != sorted(keys):
            add("ERROR", "Works Cited entries aren't in alphabetical order.")
        for p in entries:
            if re.search(r"https?://", p.text) and re.search(r"https?://\S+", p.text) and "http" in p.text.lower():
                add("WARN", "MLA 9 drops https:// from URLs. Start at www. or the domain.")
                break

    return issues


def main():
    if len(sys.argv) < 2:
        print("usage: python mla_checker.py essay.docx")
        return
    issues = check_mla(sys.argv[1])
    if not issues:
        print("No MLA problems found. Nice.")
        return
    errors = [m for l, m in issues if l == "ERROR"]
    warns = [m for l, m in issues if l == "WARN"]
    print(f"\n{len(errors)} problems, {len(warns)} warnings\n")
    for m in errors:
        print(f"  [X] {m}")
    for m in warns:
        print(f"  [!] {m}")


if __name__ == "__main__":
    main()
