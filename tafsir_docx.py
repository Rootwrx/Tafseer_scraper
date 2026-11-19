"""
tafsir_docx.py — build the Word (.docx) files from the raw data saved by tafsir_scraper.py.

  build_outputs()   one source  → per-surah files, or one merged file (--merge)
  build_combined()  several sources side by side → per-surah files, or one merged file
  configure()       the look: line spacing, tafsir font, page colour, verse colour, ligatures …
                    (called once by tafsir_scraper.main() before anything is built)
  SOURCE_INTRO      {slug: intro text}, filled by tafsir_scraper.main() for merged files

Needs python-docx (pip install python-docx).  No network access happens in this module.
"""

import re
import time
from pathlib import Path

try:
    import docx as _docxlib
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Cm, Pt, RGBColor
except ImportError:
    _docxlib = None

from tafsir_common import (
    BASE, SUWAR, _TAG_RE, ar_nums, combined_dir, load_records, log, quran_heading,
    raw_paths, source_info, src_dir, src_title, sub_dir, surah_file, surahs_stem,
)

# ═══════════════════════════════════════════════════════════════════════════
# TEXT CLEANING  (mirrors the site's prep_data: footnotes [[…]], (p-N) markers)
# ═══════════════════════════════════════════════════════════════════════════

_PAGE_MARK_RE = re.compile(r"\(p-[\d٠-٩]+\)")
_FN_RE = re.compile(r" ?\[\[([\s\S]*?\]*)\]\]")


# ─── Arabic honorific ligatures (U+FD40–FD4F etc.) ─────────────────────────
# Most Word fonts (Traditional Arabic, …) have no glyph for these, so they show as boxes.
# They are replaced by the spelled-out phrase.  ﷺ (U+FDFA) is common in fonts, so by
# default it is kept; use --ligatures expand to spell it out too.
LIGATURES = "auto"   # set from --ligatures: auto | expand | keep
_LIG_COMMON = {
    "\uFD40": "رحمه الله", "\uFD41": "رضي الله عنه", "\uFD42": "رضي الله عنها",
    "\uFD43": "رضي الله عنهم", "\uFD44": "رضي الله عنهما", "\uFD45": "رضي الله عنهن",
    "\uFD46": "صلى الله عليه وآله", "\uFD47": "عليه السلام", "\uFD48": "عليهم السلام",
    "\uFD49": "عليهما السلام", "\uFD4A": "عليه الصلاة والسلام", "\uFD4B": "قدس سره",
    "\uFD4C": "صلى الله عليه وعلى آله وسلم", "\uFD4D": "عليها السلام",
    "\uFD4E": "تبارك وتعالى", "\uFD4F": "رحمهم الله",
    "\uFDFB": "جل جلاله", "\uFDFE": "سبحانه وتعالى", "\uFDFF": "عز وجل",
}
_LIG_ALL = {**_LIG_COMMON, "\uFDFA": "صلى الله عليه وسلم", "\uFDF2": "الله", "\uFDF3": "أكبر",
            "\uFDF4": "محمد", "\uFDF5": "سلام", "\uFDF6": "رسول", "\uFDF7": "عليه",
            "\uFDF8": "وسلم", "\uFDF9": "صلى"}


def expand_ligatures(text: str) -> str:
    table = {"keep": {}, "expand": _LIG_ALL}.get(LIGATURES, _LIG_COMMON)
    return "".join(table.get(ch, ch) for ch in text) if table else text


BASMALA = "بِسْمِ اللَّهِ الرَّحْمَنِ الرَّحِيمِ"
BASMALA_MODE = "glyph"   # "glyph" = keep the ﷽ ligature (only if the font has it; Kitab may not → use "text") | "text" = spell it out
BASMALA_LINES = (BASMALA, "\uFDFD")
# a line made only of asterisks ("* * *") is a section separator: centred + red in DOCX
_SEP_LINE_RE = re.compile(r"^(?:\*\s*){3,}$")
# poetry: the site marks each verse line with U+060E (؎) and separates the two hemistichs with
# "∗∗∗"; the site shows them centred + red.  The ؎ sign renders as junk in most fonts -> dropped.
POEM_MARK = "\u060E"
POEM_COLOR = "C00000"
_HEMI_RE = re.compile(r"\s*[\u2217*]{3}\s*")


def is_poem_line(line: str) -> bool:
    return line.lstrip().startswith(POEM_MARK)


def poem_text(line: str) -> str:
    t = line.strip().lstrip(POEM_MARK).strip()
    return _HEMI_RE.sub("\u2003\u2003\u2003", t)   # em-spaces between the two hemistichs
# the site writes it as <span class="basmalah">﷽</span>; the single ligature glyph ﷽ (U+FDFD)
# is missing/truncated in most Word fonts, so it is spelled out on its own line
_BASMALA_RE = re.compile(r'<span[^>]*basmal[^>]*>.*?</span>|\uFDFD', re.S | re.I)


def clean_text(raw: str) -> tuple[str, list[str]]:
    """Return (body_text, footnotes).  Footnotes become (١) markers in the body."""
    if not raw:
        return "", []
    t = raw if LIGATURES == "keep" else _BASMALA_RE.sub(
        "\n" + (BASMALA if BASMALA_MODE == "text" else "\uFDFD") + "\n", raw)
    t = expand_ligatures(t)
    t = re.sub(r"<br\s*/?>", "\n", t)
    t = _TAG_RE.sub("", t)
    t = _PAGE_MARK_RE.sub("", t)
    notes: list[str] = []

    def _fn(m):
        notes.append(m.group(1).replace("<br>", "\n").strip())
        return f" ({ar_nums(len(notes))})"

    t = _FN_RE.sub(_fn, t)
    t = re.sub(r"[ \t]+\n", "\n", t)
    t = re.sub(r"\n{3,}", "\n\n", t).strip()
    return t, notes


# ─── red Quran quotations inside the tafsir text ──────────────────────────
VERSE_COLOR: str | None = "0F6B9C"   # set from --verse_color (None = off)
VERSE_HEURISTIC = True               # set from --verse_heuristic

# Explicit verse brackets: ﴿ ﴾ (U+FD3E/FD3F, any order), ornate parens ❨ ❩ ❪ ❫, and { }.
_DELIMS = "\uFD3E\uFD3F\u2768-\u276B{}"
_PRIMARY_RE = re.compile(f"([{_DELIMS}][^{_DELIMS}]*[{_DELIMS}])")
# Fallback: text inside ( ) « » [ ] that is heavily vowelled is a Quran quote
# (commentary is mostly unvowelled; refs like [النساء ١١٣] and notes like (١) have no vowels).
_HEUR_RE = re.compile(r"(\([^()]+\)|«[^«»]+»|\[[^\[\]]+\])")
_AR_LETTER = re.compile("[\u0621-\u064A\u0671]")
_AR_MARK = re.compile("[\u064B-\u0652\u0670]")


def _looks_vowelled(seg: str) -> bool:
    letters = len(_AR_LETTER.findall(seg))
    return letters >= 3 and len(_AR_MARK.findall(seg)) / letters >= 0.6


def _heuristic_split(text: str) -> list[tuple[str, bool]]:
    if not VERSE_HEURISTIC:
        return [(text, False)]
    out = []
    for i, seg in enumerate(_HEUR_RE.split(text)):
        if seg:
            out.append((seg, bool(i % 2) and _looks_vowelled(seg)))
    return out


def split_verses(text: str) -> list[tuple[str, bool]]:
    """'a ﴿x﴾ b' -> [('a ', False), ('﴿x﴾', True), (' b', False)]"""
    out: list[tuple[str, bool]] = []
    for i, seg in enumerate(_PRIMARY_RE.split(text)):
        if not seg:
            continue
        if i % 2:
            out.append((seg, True))
        else:
            out.extend(_heuristic_split(seg))
    # merge neighbours with the same flag
    merged: list[tuple[str, bool]] = []
    for seg, flag in out:
        if merged and merged[-1][1] == flag:
            merged[-1] = (merged[-1][0] + seg, flag)
        else:
            merged.append((seg, flag))
    return merged


def build_outputs(out_dir: Path, src: str, surahs: list[int], args, quran) -> None:
    """DOCX output:
         (default)   one file per surah   -> <src>/docx/<surah>.docx
         --merge     one file for all     -> <src>/<name>.docx"""
    base = src_dir(out_dir, src)
    title = src_title(src)
    sections = []
    for s in surahs:
        recs = load_records(raw_paths(out_dir, src, s)[0])
        if not recs:
            continue
        # drop duplicate groups (resume landing mid-group can re-fetch the same block)
        uniq = {(r["a_start"], r["a_end"]): r for r in recs}
        sections.append((s, sorted(uniq.values(), key=lambda r: r["a_start"])))
    if not sections:
        log(f"  [!] {title}: no data saved for the requested surah(s) — this source has no "
            f"material for them (or they were not downloaded), so no files were built")
        return
    stem = surahs_stem([s for s, _ in sections])   # name reflects what the file really contains

    def target(fmt: str, ext: str, s: int | None = None) -> Path:
        if not args.merge and s is not None:
            return sub_dir(base, fmt) / f"{surah_file(s)}.{ext}"
        base.mkdir(parents=True, exist_ok=True)
        return base / f"{stem}.{ext}"

    if args.docx:
        if not args.merge:
            for s, recs in sections:
                build_docx(target("docx", "docx", s), src, [(s, recs)], quran,
                           args.font, args.docx_size, title_extra=f"سورة {SUWAR[s - 1][1]}")
        else:
            build_docx(target("docx", "docx"), src, sections, quran, args.font, args.docx_size,
                       title_extra="" if len(sections) == 1 else stem,
                       intro=SOURCE_INTRO.get(src, ""))


# ═══════════════════════════════════════════════════════════════════════════
# DOCX  (python-docx — RTL, real Heading styles, footnotes, page numbers)
# ═══════════════════════════════════════════════════════════════════════════

_DOCX_COLORS = {"title": "805830", "surah": "256080", "ayah": "B8451A",
                "quran": "1A7A1A", "note": "555555", "ink": "1A1008", "sep": "C00000"}

# <w:pPr> children that must come AFTER <w:bidi> (schema order matters in Word)
_PPR_AFTER_BIDI = (
    "w:adjustRightInd", "w:snapToGrid", "w:spacing", "w:ind", "w:contextualSpacing",
    "w:mirrorIndents", "w:suppressOverlap", "w:jc", "w:textDirection", "w:textAlignment",
    "w:textboxTightWrap", "w:outlineLvl", "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr",
    "w:pPrChange",
)
_RPR_AFTER_BCS = (
    "w:i", "w:iCs", "w:caps", "w:smallCaps", "w:strike", "w:dstrike", "w:outline",
    "w:shadow", "w:emboss", "w:imprint", "w:noProof", "w:snapToGrid", "w:vanish",
    "w:webHidden", "w:color", "w:spacing", "w:w", "w:kern", "w:position", "w:sz",
    "w:szCs", "w:highlight", "w:u", "w:effect", "w:bdr", "w:shd", "w:fitText",
    "w:vertAlign", "w:rtl", "w:cs", "w:em", "w:lang", "w:eastAsianLayout",
    "w:specVanish", "w:oMath",
)
_RPR_AFTER_SZCS = _RPR_AFTER_BCS[_RPR_AFTER_BCS.index("w:szCs") + 1:]
_RPR_AFTER_ICS = _RPR_AFTER_BCS[_RPR_AFTER_BCS.index("w:iCs") + 1:]


def _d_bidi(par) -> None:
    """Mark a paragraph (or style) right-to-left."""
    pf = par.paragraph_format
    pf.left_indent = pf.right_indent = pf.first_line_indent = Pt(0)
    ppr = (par._p if hasattr(par, "_p") else par.element).get_or_add_pPr()
    if ppr.find(qn("w:bidi")) is None:
        ppr.insert_element_before(OxmlElement("w:bidi"), *_PPR_AFTER_BIDI)


def _d_run(par, text: str, font: str, size: float, bold=False, color: str | None = None):
    r = par.add_run(text)
    r.font.rtl = True
    r.font.complex_script = True
    rpr = r._r.get_or_add_rPr()
    rf = rpr.get_or_add_rFonts()
    for k in ("w:cs", "w:ascii", "w:hAnsi"):
        rf.set(qn(k), font)
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.italic = False
    szcs = OxmlElement("w:szCs")                      # Word sizes Arabic from szCs
    szcs.set(qn("w:val"), str(int(round(size * 2))))
    rpr.insert_element_before(szcs, *_RPR_AFTER_SZCS)
    bcs = OxmlElement("w:bCs")                        # ...and weights it from bCs
    if not bold:
        bcs.set(qn("w:val"), "0")
    rpr.insert_element_before(bcs, *_RPR_AFTER_BCS)
    ics = OxmlElement("w:iCs")                        # Arabic italics come from iCs (Heading 4 style is italic)
    ics.set(qn("w:val"), "0")
    rpr.insert_element_before(ics, *_RPR_AFTER_ICS)
    if color:
        r.font.color.rgb = RGBColor.from_string(color)
    return r


def _d_par(doc, text: str, font: str, size: float, *, bold=False, color=None,
           align=None, before=0, after=4, style=None):
    par = doc.add_paragraph(style=style) if style else doc.add_paragraph()
    _d_bidi(par)
    par.alignment = align if align is not None else WD_ALIGN_PARAGRAPH.LEFT  # LEFT == "start" in RTL
    par.paragraph_format.space_before = Pt(before)
    par.paragraph_format.space_after = Pt(after)
    if text:
        _d_run(par, text, font, size, bold, color)
    return par


def _d_heading(doc, text: str, level: int, font: str, size: float, color: str, page_break=False,
               bold: bool = True):
    # *doc* may be a Document or a table cell (cells have no add_heading)
    h = (doc.add_heading(level=level) if hasattr(doc, "add_heading")
         else doc.add_paragraph(style=f"Heading {level}"))
    for r in list(h.runs):
        r.text = ""
    _d_bidi(h)
    h.alignment = WD_ALIGN_PARAGRAPH.LEFT
    h.paragraph_format.space_before = Pt(14 if level == 1 else 8)
    h.paragraph_format.space_after = Pt(4)
    h.paragraph_format.keep_with_next = True
    if page_break:
        h.paragraph_format.page_break_before = True
    _d_run(h, text, font, size, bold, color)
    return h


def _d_page_number_footer(doc, font: str) -> None:
    par = doc.sections[0].footer.paragraphs[0]
    par.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _d_bidi(par)
    _d_run(par, "— ", font, 10, color="888888")
    r = par.add_run()
    for kind, txt in (("begin", None), (None, "PAGE"), ("end", None)):
        if kind:
            el = OxmlElement("w:fldChar")
            el.set(qn("w:fldCharType"), kind)
        else:
            el = OxmlElement("w:instrText")
            el.set(qn("xml:space"), "preserve")
            el.text = txt
        r._r.append(el)
    _d_run(par, " —", font, 10, color="888888")


LINE_SPACING = 0.9   # set from --line_spacing
VERSE_SIZE_DELTA = 2   # pt smaller than the tafsir text for verse quotations inside it (0 = same size)
TAFSIR_FONT = "KFGQPC Uthman Taha Naskh"   # set from --tafsir_font: tafsir names + tafsir text (verses stay in --font)
PAGE_BG: str | None = "F3EEE2"   # set from --page_bg (None = white page)


def configure(*, line_spacing, tafsir_font, page_bg, ligatures, verse_heuristic, verse_color) -> None:
    """Set the look of the Word files (values come straight from the command-line arguments)."""
    global VERSE_COLOR, VERSE_HEURISTIC, LIGATURES, PAGE_BG, LINE_SPACING, TAFSIR_FONT
    LINE_SPACING = line_spacing
    TAFSIR_FONT = tafsir_font.split(",")[0].strip() or "KFGQPC Uthman Taha Naskh"
    PAGE_BG = None if page_bg.lower() in ("none", "off", "white", "") else page_bg.lstrip("#").upper()
    LIGATURES = ligatures
    VERSE_HEURISTIC = verse_heuristic
    VERSE_COLOR = None if verse_color.lower() in ("none", "off", "") else verse_color.lstrip("#").upper()


def _d_page_bg(doc, hex_color: str | None) -> None:
    """Colour the whole page (Word: Design ▸ Page Color).  Printing it needs Word's
    'Print background colors' option; PDF export from Word keeps it."""
    if not hex_color:
        return
    body = doc.element
    for old in body.findall(qn("w:background")):
        body.remove(old)
    bg = OxmlElement("w:background")
    bg.set(qn("w:color"), hex_color)
    body.insert(0, bg)                       # must be the first child of <w:document>
    st = doc.settings.element
    if st.find(qn("w:displayBackgroundShape")) is None:
        el = OxmlElement("w:displayBackgroundShape")
        zoom = st.find(qn("w:zoom"))
        if zoom is not None:
            zoom.addnext(el)
        else:
            st.insert(0, el)


def _d_setup(doc, font: str, size: float) -> None:
    _d_page_bg(doc, PAGE_BG)
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.top_margin = Cm(1.8)
    sec.bottom_margin = Cm(1.3)      # smaller bottom margin = more lines per page
    sec.footer_distance = Cm(0.5)    # page number sits close to the bottom edge
    sec.left_margin = sec.right_margin = Cm(1.8)
    # section-level RTL
    sp = sec._sectPr
    if sp.find(qn("w:bidi")) is None:
        sp.insert_element_before(OxmlElement("w:bidi"), "w:rtlGutter", "w:docGrid",
                                 "w:printerSettings", "w:sectPrChange")
    normal = doc.styles["Normal"]
    normal.font.name = font
    normal.font.size = Pt(size)
    normal.paragraph_format.line_spacing = LINE_SPACING
    rf = normal.element.get_or_add_rPr().get_or_add_rFonts()
    rf.set(qn("w:cs"), font)
    _d_bidi(normal)
    for name in ("Title", "Heading 1", "Heading 2", "Heading 3", "Heading 4"):
        hs = doc.styles[name]   # built-in headings default to LTR + Calibri + italics
        hs.font.italic = False
        hs.paragraph_format.line_spacing = LINE_SPACING
        hs.element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:cs"), font)
        hs.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        _d_bidi(hs)
    _d_page_number_footer(doc, font)


def _d_info_lines(doc, src: str, font: str, size: float) -> None:
    """Cover page: one centred 'المؤلف: …' / 'العقيدة: …' line per field the source has."""
    C = _DOCX_COLORS
    for label, val in source_info(src):
        par = _d_par(doc, "", font, size, align=WD_ALIGN_PARAGRAPH.CENTER, after=2)
        _d_run(par, f"{label}: ", font, size, color=C["note"])
        _d_run(par, val, font, size, color=C["ink"])


# ── book intro (get_info.php) — shown once on the cover side of MERGED docx files only ──
SOURCE_INTRO: dict[str, str] = {}   # slug → cleaned intro text, filled in main() before building


def _d_intro(doc, text: str, font: str, tf: str, size: float, level: int) -> None:
    """Intro text: '* title' lines become headings (*level*), everything else is a paragraph."""
    C = _DOCX_COLORS
    for raw in text.split("\n"):
        line = raw.strip()
        if not line:
            continue
        if line.startswith("*"):
            head = line.lstrip("*").strip()
            if head:
                _d_heading(doc, head, level, font, size, C["surah"], bold=True)
            continue
        _d_rich(doc, line, tf, size, C["ink"], verse_font=font, align=WD_ALIGN_PARAGRAPH.JUSTIFY)


BASMALA_ORNAMENT = "\uFDFD"          # ﷽
BASMALA_FONT = "Kitab"               # font used for the ﷽ line in docx files
BASMALA_SKIP_SURAHS = {9}            # surahs that get no ﷽ (At-Tawba).  Add 1 here to skip Al-Fatiha, whose ayah 1 IS the basmala


def _d_surah_basmala(doc, surah: int, first_ayah: int, size: float) -> None:
    """﷽ centred under the surah heading, before the first ayah (not for surah 9, and only
    when the content really starts at ayah 1 — a partial --ayahs range gets none)."""
    if surah in BASMALA_SKIP_SURAHS or first_ayah != 1:
        return
    par = _d_par(doc, BASMALA_ORNAMENT, BASMALA_FONT, size + 10, color="C99200",
                 align=WD_ALIGN_PARAGRAPH.CENTER, before=4, after=8)
    par.paragraph_format.keep_with_next = True


def _d_sections(doc, src: str, sections, quran, font: str, size: float, page_break_surah: bool,
                break_first: bool = False):
    """Emit surah → ayah-group content into *doc*.  sections = [(surah, records)]."""
    C = _DOCX_COLORS
    for idx, (s, recs) in enumerate(sections):
        _d_heading(doc, f"سورة {SUWAR[s - 1][1]}", 1, font, size + 6, C["surah"],
                   page_break=page_break_surah and (idx > 0 or break_first))
        if recs:
            _d_surah_basmala(doc, s, recs[0]["a_start"], size)
        for r in recs:
            rng = (ar_nums(r["a_start"]) if r["a_start"] == r["a_end"]
                   else f"{ar_nums(r['a_start'])}-{ar_nums(r['a_end'])}")
            _d_heading(doc, f"الآية {rng}", 2, font, size + 2, C["ayah"], bold=False)
            head = quran_heading(quran, s, r["a_start"], r["a_end"])
            if head:
                _d_par(doc, head, font, size, bold=False, color=C["quran"], after=6)
            body, notes = clean_text(r["data"])
            if not body:
                _d_par(doc, "لا مواد لهذه الآية في هذا الكتاب", font, size - 2, color="888888")
                continue
            def _rich(text, sz, base_color, **kw):
                par = _d_par(doc, "", TAFSIR_FONT, sz, **kw)
                _d_text_runs(par, text, TAFSIR_FONT, font, sz, base_color)

            for line in body.split("\n"):
                if is_poem_line(line):
                    _d_par(doc, poem_text(line), TAFSIR_FONT, size, color=POEM_COLOR,
                           align=WD_ALIGN_PARAGRAPH.CENTER, before=4, after=4)
                elif line.strip() in BASMALA_LINES:
                    _rich(line.strip(), size + (2 if BASMALA_MODE == "text" else 10), "C99200", align=WD_ALIGN_PARAGRAPH.LEFT,
                          before=6, after=6)
                elif _SEP_LINE_RE.match(line.strip()):
                    _rich(line.strip(), size, C["sep"], align=WD_ALIGN_PARAGRAPH.CENTER,
                          before=6, after=6)
                elif line.strip():
                    _rich(line.strip(), size, C["ink"], align=WD_ALIGN_PARAGRAPH.JUSTIFY)
            _d_notes(doc, notes, TAFSIR_FONT, font, max(10, size - 4))


def build_docx(out_path: Path, src: str, sections, quran, font: str, size: float,
               title_extra: str = "", intro: str = "") -> None:
    if _docxlib is None:
        log("  [!] python-docx not installed — DOCX skipped (pip install python-docx)")
        return
    C = _DOCX_COLORS
    title = src_title(src)
    doc = _docxlib.Document()
    _d_setup(doc, font, size)

    # cover block
    _d_par(doc, "بِسْمِ اللَّهِ الرَّحْمَنِ الرَّحِيمِ", font, size + 2, color="C99200",
           align=WD_ALIGN_PARAGRAPH.CENTER, after=14)
    t = doc.add_heading(level=0)
    for r in list(t.runs):
        r.text = ""
    _d_bidi(t)
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _d_run(t, title, TAFSIR_FONT, size + 14, True, C["title"])
    if title_extra:
        _d_par(doc, title_extra, font, size + 2, color=C["surah"], align=WD_ALIGN_PARAGRAPH.CENTER)
    _d_info_lines(doc, src, TAFSIR_FONT, size)
    _d_par(doc, f"المصدر: {BASE}/{src} | تاريخ التحميل: {time.strftime('%Y-%m-%d')}",
           font, 10, color="888888", align=WD_ALIGN_PARAGRAPH.CENTER, after=12)

    if intro:    # merged files only (the caller passes "" for per-surah files)
        _d_heading(doc, "نبذة عن الكتاب", 1, font, size + 6, C["title"])
        _d_intro(doc, intro, font, TAFSIR_FONT, size, 2)
    _d_sections(doc, src, sections, quran, font, size, page_break_surah=True,
                break_first=bool(intro))
    doc.save(str(out_path))
    n = sum(len(r) for _, r in sections)
    log(f"  [docx] saved → {out_path}  ({len(sections)} surah(s), {n} entries)")


def combined_entries(out_dir: Path, order: list[str], s: int) -> list[dict]:
    """Merge the saved records of several sources for one surah.

    Returns [{"a_start", "a_end", "items": [(idx, src, rec)], "notes": [(idx, src, rec)]}]
    in ayah order.  An entry starts at every ayah where at least one source has a block
    beginning; `items` holds those blocks (in --combine order, idx = position in that
    order, 1-based); `notes` holds sources whose earlier, longer block already covers it.
    """
    per_src = {}
    for src in order:
        recs = load_records(raw_paths(out_dir, src, s)[0])
        uniq = {(r["a_start"], r["a_end"]): r for r in recs if clean_text(r["data"])[0]}
        per_src[src] = sorted(uniq.values(), key=lambda r: (r["a_start"], r["a_end"]))
    starts = sorted({r["a_start"] for rs in per_src.values() for r in rs})
    entries = []
    for a in starts:
        items, notes = [], []
        for idx, src in enumerate(order, 1):
            starting = [r for r in per_src[src] if r["a_start"] == a]
            if starting:
                items.append((idx, src, max(starting, key=lambda r: r["a_end"])))
                continue
            cover = [r for r in per_src[src] if r["a_start"] < a <= r["a_end"]]
            if cover:
                notes.append((idx, src, cover[0]))
        entries.append({"a_start": a, "a_end": max(r["a_end"] for _, _, r in items),
                        "items": items, "notes": notes})
    return entries


def _rng(a: int, b: int) -> str:
    return ar_nums(a) if a == b else f"{ar_nums(a)}-{ar_nums(b)}"


def _src_label(idx: int, src: str) -> str:
    return src_title(src)


def _see_before(idx: int, src: str, rec: dict) -> str:
    return (f"{_src_label(idx, src)}: تقدّم الكلام في الآيات "
            f"{_rng(rec['a_start'], rec['a_end'])}")


_NUM_MARK_RE = re.compile(r"(\(\s*[\d٠-٩]+\s*\))")   # (١)  (٢)  (12) ... footnote marks & numbering


def _d_text_runs(par, text: str, text_font: str, verse_font: str, sz: float, base_color: str,
                 bold: bool = False) -> None:
    """Add *text* to *par*: Quran quotations in the verse font/colour, and every (١)-style
    number in the verse colour too; everything else in *base_color*."""
    for seg, is_verse in split_verses(text):
        if is_verse:
            _d_run(par, seg, verse_font, sz - VERSE_SIZE_DELTA, bold=False,
                   color=VERSE_COLOR or base_color)
            continue
        for part in _NUM_MARK_RE.split(seg):
            if not part:
                continue
            is_num = bool(_NUM_MARK_RE.fullmatch(part))
            _d_run(par, part, text_font, sz, bold=bold,
                   color=_note_color() if is_num else base_color)


def _d_rich(doc, text: str, font: str, sz: float, base_color: str, bold_text: bool = False,
            verse_font: str | None = None, **kw):
    """Text in *font*; Quran quotations use *verse_font* (defaults to *font*)."""
    par = _d_par(doc, "", font, sz, **kw)
    _d_text_runs(par, text, font, verse_font or font, sz, base_color, bold_text)
    return par


NOTE_RULE_CM = 6.0     # length of the line above the footnotes
NOTE_RULE_SIDE = "end"   # logical end side (= left in RTL) is indented, so the line stays at the right


def _note_color() -> str:
    """Footnote numbers + the line above them use the verse colour (blue by default)."""
    return VERSE_COLOR or "0F6B9C"


def _d_note_rule(container):
    """Short coloured line above the footnotes, at the right-hand (start) side like in printed books."""
    par = _d_par(container, "", TAFSIR_FONT, 4, before=8, after=4)
    ppr = par._p.get_or_add_pPr()
    bdr = OxmlElement("w:pBdr")
    b = OxmlElement("w:bottom")
    b.set(qn("w:val"), "single")
    b.set(qn("w:sz"), "8")
    b.set(qn("w:space"), "1")
    b.set(qn("w:color"), _note_color())
    bdr.append(b)
    ppr.insert_element_before(
        bdr, "w:shd", "w:tabs", "w:suppressAutoHyphens", "w:kinsoku", "w:wordWrap",
        "w:overflowPunct", "w:topLinePunct", "w:autoSpaceDE", "w:autoSpaceDN", "w:bidi",
        *_PPR_AFTER_BIDI)
    # indent from the far (left) side so only NOTE_RULE_CM of the text width is underlined
    text_w = 21.0 - 2 * 1.8
    ind = OxmlElement("w:ind")
    ind.set(qn(f"w:{NOTE_RULE_SIDE}"), str(int((text_w - NOTE_RULE_CM) / 2.54 * 1440)))
    for old in ppr.findall(qn("w:ind")):
        ppr.remove(old)
    ppr.insert_element_before(ind, "w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap",
                              "w:jc", "w:textDirection", "w:textAlignment", "w:textboxTightWrap",
                              "w:outlineLvl", "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr",
                              "w:pPrChange")
    return par


def _d_notes(container, notes: list[str], text_font: str, verse_font: str, sz: float) -> list:
    """Footnotes: a coloured rule, then «(١) text» with the number in the verse colour."""
    if not notes:
        return []
    out = [_d_note_rule(container)]
    for i, n in enumerate(notes, 1):
        par = _d_par(container, "", text_font, sz, after=2)
        _d_run(par, f"({ar_nums(i)}) ", text_font, sz, color=_note_color())
        _d_text_runs(par, n, text_font, verse_font, sz, _DOCX_COLORS["note"])
        out.append(par)
    return out


# light background per tafsir (cycled in --combine order)
SRC_BG = ["FFF4CC", "DDEEFF", "DFF3E0", "FBE0EA", "EBE0F7", "FFE8D1", "D9F4F6", "EEF5D5"]
# darker tone of the same hue, used for that tafsir's title
SRC_FG = ["8A6A00", "1F5E9E", "2E7D32", "A62A57", "5E3A99", "B5560B", "0F7480", "5F7A12"]


def _d_shade(par, fill: str) -> None:
    """Paragraph background (w:shd must sit before w:bidi/w:tabs/... in <w:pPr>)."""
    ppr = par._p.get_or_add_pPr()
    for old in ppr.findall(qn("w:shd")):
        ppr.remove(old)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    # same-colour box border = inner padding; identical on every paragraph of a card so
    # Word merges them into ONE rounded-looking block instead of striped lines
    bdr = OxmlElement("w:pBdr")
    for side in ("top", "left", "bottom", "right"):
        b = OxmlElement(f"w:{side}")
        b.set(qn("w:val"), "single")
        b.set(qn("w:sz"), "4")
        b.set(qn("w:space"), "6")
        b.set(qn("w:color"), fill)
        bdr.append(b)
    for old in ppr.findall(qn("w:pBdr")):
        ppr.remove(old)
    ppr.insert_element_before(
        bdr, "w:shd", "w:tabs", "w:suppressAutoHyphens", "w:kinsoku", "w:wordWrap",
        "w:overflowPunct", "w:topLinePunct", "w:autoSpaceDE", "w:autoSpaceDN", "w:bidi",
        *_PPR_AFTER_BIDI)
    ppr.insert_element_before(
        shd, "w:tabs", "w:suppressAutoHyphens", "w:kinsoku", "w:wordWrap", "w:overflowPunct",
        "w:topLinePunct", "w:autoSpaceDE", "w:autoSpaceDN", "w:bidi", *_PPR_AFTER_BIDI)


def _d_outline(par, level: int) -> None:
    """Give a normal paragraph an outline level (0 = level 1) so Word shows a collapse arrow
    on it and hides everything below it, up to the next paragraph of the same/higher level."""
    ppr = par._p.get_or_add_pPr()
    for old in ppr.findall(qn("w:outlineLvl")):
        ppr.remove(old)
    el = OxmlElement("w:outlineLvl")
    el.set(qn("w:val"), str(level))
    ppr.insert_element_before(el, "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr", "w:pPrChange")


def _d_tafsir_block(container, item, e, font: str, size: float, title_size: float = 20,
                    bold: bool = False) -> None:
    """One tafsir (title + text + footnotes) into a Document or a table cell."""
    C = _DOCX_COLORS
    idx, src, rec = item
    body, notes = clean_text(rec["data"])
    label = _src_label(idx, src)
    if (rec["a_start"], rec["a_end"]) != (e["a_start"], e["a_end"]):
        label += f" — الآيات {_rng(rec['a_start'], rec['a_end'])}"
    tf = TAFSIR_FONT   # tafsir name + text; verses stay in *font* (Kitab)
    block = [_d_heading(container, label, 4, tf, title_size, SRC_FG[(idx - 1) % len(SRC_FG)],
                        bold=False)]
    for line in body.split("\n"):
        if is_poem_line(line):
            block.append(_d_par(container, poem_text(line), tf, size, color=POEM_COLOR,
                                align=WD_ALIGN_PARAGRAPH.CENTER, before=4, after=4))
        elif line.strip() in BASMALA_LINES:
            block.append(_d_rich(container, line.strip(), tf,
                                 size + (2 if BASMALA_MODE == "text" else 10), "C99200",
                                 verse_font=font,
                                 align=WD_ALIGN_PARAGRAPH.LEFT, before=6, after=6))
        elif _SEP_LINE_RE.match(line.strip()):
            block.append(_d_rich(container, line.strip(), tf, size, C["sep"], verse_font=font,
                                 align=WD_ALIGN_PARAGRAPH.CENTER, before=6, after=6))
        elif line.strip():
            block.append(_d_rich(container, line.strip(), tf, size, C["ink"], bold_text=bold,
                                 verse_font=font, align=WD_ALIGN_PARAGRAPH.JUSTIFY))
    block.extend(_d_notes(container, notes, tf, font, max(10, size - 4)))
    block[0].paragraph_format.space_before = Pt(10)
    block[-1].paragraph_format.space_after = Pt(10)


def build_combined_docx(out_path: Path, order: list[str], sections, quran, font: str,
                        size: float, title_extra: str = "", title_size: float = 20,
                        bold: bool = False, intros: dict | None = None) -> None:
    """sections = [(surah, entries)] from combined_entries().
    intros = {slug: text} → each book's intro once, after the cover (merged files only)."""
    if _docxlib is None:
        log("  [!] python-docx not installed — DOCX skipped (pip install python-docx)")
        return
    C = _DOCX_COLORS
    doc = _docxlib.Document()
    _d_setup(doc, font, size)
    _d_par(doc, "بِسْمِ اللَّهِ الرَّحْمَنِ الرَّحِيمِ", font, size + 2, color="C99200",
           align=WD_ALIGN_PARAGRAPH.CENTER, after=14)
    t = doc.add_heading(level=0)
    for r in list(t.runs):
        r.text = ""
    _d_bidi(t)
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _d_run(t, "جمع التفاسير", font, size + 14, True, C["title"])
    if title_extra:
        _d_par(doc, title_extra, font, size + 2, color=C["surah"], align=WD_ALIGN_PARAGRAPH.CENTER)
    for idx, src in enumerate(order, 1):
        _d_par(doc, _src_label(idx, src), TAFSIR_FONT, title_size, color=SRC_FG[(idx - 1) % len(SRC_FG)],
               align=WD_ALIGN_PARAGRAPH.CENTER, after=0)
        info = "  |  ".join(f"{l}: {v}" for l, v in source_info(src))
        if info:
            _d_par(doc, info, TAFSIR_FONT, max(10, size - 4), color="888888",
                   align=WD_ALIGN_PARAGRAPH.CENTER, after=4)
    _d_par(doc, f"تاريخ التحميل: {time.strftime('%Y-%m-%d')}", font, 10, color="888888",
           align=WD_ALIGN_PARAGRAPH.CENTER, before=8, after=12)

    intros = {s_: t_ for s_, t_ in (intros or {}).items() if t_}
    if intros:
        _d_heading(doc, "نبذة عن الكتب", 1, font, size + 6, C["title"])
        for idx, src in enumerate(order, 1):
            if src in intros:
                _d_heading(doc, _src_label(idx, src), 2, TAFSIR_FONT, title_size,
                           SRC_FG[(idx - 1) % len(SRC_FG)], bold=False)
                _d_intro(doc, intros[src], font, TAFSIR_FONT, size, 3)

    for k, (s, entries) in enumerate(sections):
        _d_heading(doc, f"سورة {SUWAR[s - 1][1]}", 1, font, size + 6, C["surah"],
                   page_break=k > 0 or bool(intros))
        if entries:
            _d_surah_basmala(doc, s, entries[0]["a_start"], size)
        for e in entries:
            _d_heading(doc, f"الآية {_rng(e['a_start'], e['a_end'])}", 2, font, size + 2, C["ayah"], bold=False)
            head = quran_heading(quran, s, e["a_start"], e["a_end"])
            if head:
                vp = _d_par(doc, head, font, size, bold=False, color=C["quran"], after=6)
                _d_outline(vp, 2)   # collapsible: hides all tafsirs of this ayah
            for it in e["items"]:
                _d_tafsir_block(doc, it, e, font, size, title_size, bold)
            for idx, src, rec in e["notes"]:
                _d_par(doc, _see_before(idx, src, rec), font, max(10, size - 4),
                       color="888888", after=10)
    doc.save(str(out_path))
    log(f"  [docx] saved → {out_path}  ({len(sections)} surah(s), {len(order)} source(s))")


def build_combined(out_dir: Path, order: list[str], surahs: list[int], args, quran) -> None:
    root = combined_dir(out_dir, order)
    sections = []
    for s in surahs:
        entries = combined_entries(out_dir, order, s)
        if entries:
            sections.append((s, entries))
    if not sections:
        log("[!] combine: no data found for the requested sources/surahs")
        return
    stem = surahs_stem([s for s, _ in sections])

    # same rule as build_outputs: one file per surah by default, --merge = one file for all
    if args.docx:
        if not args.merge:
            d = sub_dir(root, "docx")
            for s, entries in sections:
                build_combined_docx(d / f"{surah_file(s)}.docx", order, [(s, entries)],
                                    quran, args.font, args.docx_size,
                                    title_extra=f"سورة {SUWAR[s - 1][1]}",
                                    title_size=args.tafsir_title_size,
                                    bold=args.tafsir_bold)
        else:
            build_combined_docx(root / f"{stem}.docx", order, sections, quran,
                                args.font, args.docx_size,
                                title_extra="" if len(sections) == 1 else stem,
                                title_size=args.tafsir_title_size, bold=args.tafsir_bold,
                                intros={s_: SOURCE_INTRO.get(s_, "") for s_ in order})
