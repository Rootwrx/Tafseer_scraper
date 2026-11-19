#!/usr/bin/env python3
"""
iraab_graphs_docx.py — build Word files from the tafsir.app "الإعراب المرسوم" (iraab-graphs) data.

Input  : the raw JSONL files written by tafsir_scraper.py, e.g.
           tafsir_output/.../الإعراب المرسوم/raw/surah_004.jsonl
         one record per line:
           {"src": "iraab-graphs", "s": 4, "a_start": 24, "a_end": 24, "data": "<svg ...>...</svg>"}
         Each `data` is a self-contained SVG (fonts embedded as base64 @font-face) that holds
         one or more <g data-graph="N"> groups = one syntax graph per sentence.

Why not put the SVG in Word directly?
         Word ignores fonts embedded inside an SVG, so Arabic would fall back to another
         font and lose its shaping.  Instead every graph is rendered in headless Chromium
         (which does load the embedded fonts) to a transparent PNG and that PNG is placed
         in the .docx.  The result looks exactly like the website.

Output : <out>/سورة النساء.docx            one file per surah        (default)
         <out>/القرآن كاملا.docx / سورة … إلى سورة … .docx   one merged file   (--merge)
         <out>/_png_cache/                  rendered graphs (re-runs are fast / resumable)

Setup  : pip install playwright python-docx pillow
         playwright install chromium

Usage
─────
  python iraab_graphs_docx.py --in_dir ./tafsir_output                 # per-surah files
  python iraab_graphs_docx.py --in_dir ./tafsir_output --merge         # + one combined file
  python iraab_graphs_docx.py --in_dir ./tafsir_output --merge --no_per_surah
  python iraab_graphs_docx.py --in_dir ./tafsir_output --surah 1-4,112-114
  python iraab_graphs_docx.py --in_dir path/to/surah_004.jsonl         # a single file works too

Optional: put this file next to tafsir_scraper.py (or give --scraper) and it reuses its surah
names and (if <out_dir>/quran.json exists) prints the ﴿ayah﴾ text above each graph set.
"""
from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import re
import sys
import time
from pathlib import Path

SRC_SLUG = "iraab-graphs"
SRC_TITLE = "الإعراب المرسوم"
SITE_BASE = "https://tafsir.app"
RE_SURAH_FILE = re.compile(r"surah_(\d{3})\.jsonl$")


# ───────────────────────── helpers ─────────────────────────
def log(*a):
    print(*a, flush=True)


def ar_nums(n) -> str:
    return str(n).translate(str.maketrans("0123456789", "٠١٢٣٤٥٦٧٨٩"))


def parse_surahs(spec: str | None) -> set[int] | None:
    if not spec or spec.lower() == "all":
        return None
    out: set[int] = set()
    for part in spec.split(","):
        part = part.strip()
        if "-" in part:
            a, b = part.split("-", 1)
            out.update(range(int(a), int(b) + 1))
        elif part:
            out.add(int(part))
    return out


def load_scraper(path: str | None):
    """Optionally import tafsir_scraper.py for SUWAR / ABS_START (surah names, quran.json index)."""
    cands = [Path(path)] if path else [Path(__file__).with_name("tafsir_scraper.py"),
                                        Path.cwd() / "tafsir_scraper.py"]
    for c in cands:
        if c.exists():
            try:
                spec = importlib.util.spec_from_file_location("tafsir_scraper", c)
                mod = importlib.util.module_from_spec(spec)
                sys.modules["tafsir_scraper"] = mod
                spec.loader.exec_module(mod)
                return mod
            except Exception as e:  # noqa: BLE001
                log(f"[!] could not import {c.name}: {e}  (continuing without it)")
    return None


def safe_name(s: str, limit: int = 120) -> str:
    return re.sub(r'[\\/:*?"<>|]+', "_", s).strip()[:limit]


def surahs_stem(surahs: list[int], name_of) -> str:
    surahs = sorted(set(surahs))
    if len(surahs) == 114:
        return "القرآن كاملا"
    if len(surahs) == 1:
        return name_of(surahs[0])
    if surahs[-1] - surahs[0] + 1 == len(surahs):
        return f"{name_of(surahs[0])} إلى {name_of(surahs[-1])}"
    runs, lo, prev = [], surahs[0], surahs[0]
    for x in surahs[1:] + [None]:
        if x is None or x != prev + 1:
            runs.append(str(lo) if lo == prev else f"{lo}-{prev}")
            lo = x
        prev = x
    return safe_name("سور " + "،".join(runs), 100)


# ───────────────────────── reading the data ─────────────────────────
def find_jsonl_files(in_path: Path) -> list[Path]:
    if in_path.is_file():
        return [in_path]
    files = sorted(p for p in in_path.rglob("surah_*.jsonl") if RE_SURAH_FILE.search(p.name))
    # keep only the iraab-graphs ones: by folder name, or by peeking at the first record
    keep = []
    for p in files:
        if SRC_SLUG in p.parts or SRC_TITLE in p.parts:
            keep.append(p)
            continue
        try:
            with open(p, encoding="utf-8") as fh:
                first = fh.readline()
            if first and json.loads(first).get("src") == SRC_SLUG:
                keep.append(p)
        except Exception:  # noqa: BLE001
            pass
    return keep


def load_records(files: list[Path], wanted: set[int] | None) -> dict[int, list[dict]]:
    """→ {surah: [records sorted by a_start]}  (duplicates by (a_start,a_end) removed, last wins)"""
    by_surah: dict[int, dict[tuple, dict]] = {}
    for f in files:
        with open(f, encoding="utf-8") as fh:
            for ln, line in enumerate(fh, 1):
                line = line.strip()
                if not line:
                    continue
                try:
                    r = json.loads(line)
                except json.JSONDecodeError:
                    log(f"[!] {f.name}:{ln} bad JSON line — skipped")
                    continue
                if r.get("src") not in (None, SRC_SLUG):
                    continue
                s = int(r.get("s", 0))
                if not s or (wanted and s not in wanted):
                    continue
                if not r.get("data") or "<svg" not in r["data"]:
                    continue
                by_surah.setdefault(s, {})[(int(r["a_start"]), int(r["a_end"]))] = r
    return {s: sorted(d.values(), key=lambda r: (int(r["a_start"]), int(r["a_end"])))
            for s, d in sorted(by_surah.items())}


# ───────────────────────── SVG → PNG ─────────────────────────
HTML_WRAP = """<!doctype html><html><head><meta charset="utf-8">
<style>html,body{{margin:0;padding:0;background:transparent}}
body{{width:{w}px}} svg{{display:block;max-width:none !important;height:auto}}</style>
</head><body>{svg}</body></html>"""


class Renderer:
    def __init__(self, cache_dir: Path, scale: float, pad: int):
        from playwright.sync_api import sync_playwright
        self.cache = cache_dir
        self.cache.mkdir(parents=True, exist_ok=True)
        self.scale, self.pad = scale, pad
        self._pw = sync_playwright().start()
        self._browser = self._pw.chromium.launch()
        self._ctx = self._browser.new_context(device_scale_factor=scale)
        self._page = self._ctx.new_page()

    def close(self):
        try:
            self._ctx.close()
            self._browser.close()
        finally:
            self._pw.stop()

    def _key(self, svg: str, idx: int) -> Path:
        h = hashlib.sha1(svg.encode("utf-8")).hexdigest()[:16]
        return self.cache / f"{h}_{self.scale:g}_{self.pad}_{idx:02d}.png"

    def render(self, svg: str) -> list[tuple[Path, int, int]]:
        """Render every <g data-graph> of the SVG as its own PNG.
        → [(png_path, css_width_px, css_height_px), ...] in graph order."""
        from PIL import Image

        # cached?  (we don't know the graph count up-front, so a ".n" marker file stores it)
        marker = self.cache / (hashlib.sha1(svg.encode("utf-8")).hexdigest()[:16]
                               + f"_{self.scale:g}_{self.pad}.n")
        if marker.exists():
            n = int(marker.read_text())
            paths = [self._key(svg, i) for i in range(n)]
            if all(p.exists() for p in paths):
                out = []
                for p in paths:
                    with Image.open(p) as im:
                        out.append((p, round(im.width / self.scale), round(im.height / self.scale)))
                return out

        m = re.search(r'\bwidth="(\d+(?:\.\d+)?)"', svg)
        width = int(float(m.group(1))) if m else 1500
        page = self._page
        page.set_viewport_size({"width": width + 40, "height": 1000})
        page.set_content(HTML_WRAP.format(w=width, svg=svg), wait_until="load")
        # make sure the embedded fonts are really loaded before we measure / screenshot
        page.evaluate("""async () => {
            const fams = ['tafsir-treebank-hafs','tafsir-treebank-labels'];
            try { await Promise.all(fams.map(f => document.fonts.load('42px "'+f+'"', 'ءابت'))); } catch(e){}
            await document.fonts.ready;
        }""")
        page.wait_for_timeout(150)

        # bounding box of every graph group (page coordinates)
        boxes = page.evaluate("""() => {
            const svg = document.querySelector('svg');
            const root = svg.getBoundingClientRect();
            const gs = [...svg.querySelectorAll(':scope > g[data-graph]')];
            const targets = gs.length ? gs : [svg];
            return targets.map(g => {
                const r = g.getBoundingClientRect();
                return {x: r.left + window.scrollX, y: r.top + window.scrollY, w: r.width, h: r.height};
            });
        }""")
        body_w = page.evaluate("document.documentElement.scrollWidth")
        body_h = page.evaluate("document.documentElement.scrollHeight")

        out: list[tuple[Path, int, int]] = []
        for i, b in enumerate(boxes):
            x0 = max(0, b["x"] - self.pad)
            y0 = max(0, b["y"] - self.pad)
            x1 = min(body_w, b["x"] + b["w"] + self.pad)
            y1 = min(body_h, b["y"] + b["h"] + self.pad)
            if x1 - x0 < 2 or y1 - y0 < 2:
                continue
            png = self._key(svg, len(out))
            page.screenshot(path=str(png), omit_background=True, full_page=True,
                            clip={"x": x0, "y": y0, "width": x1 - x0, "height": y1 - y0})
            with Image.open(png) as im:           # skip graphs that rendered fully transparent
                if im.getchannel("A").getbbox() is None:
                    png.unlink()
                    continue
            out.append((png, round(x1 - x0), round(y1 - y0)))
        marker.write_text(str(len(out)))
        return out


# ───────────────────────── DOCX ─────────────────────────
def _set_rtl_par(par):
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    pPr = par._p.get_or_add_pPr()
    if pPr.find(qn("w:bidi")) is None:
        pPr.insert(0, OxmlElement("w:bidi"))


def _font_run(run, font: str, size_pt: float, color=None, bold=False):
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Pt, RGBColor
    run.font.name = font
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    if color:
        run.font.color.rgb = RGBColor.from_string(color)
    rPr = run._r.get_or_add_rPr()
    rf = rPr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rPr.insert(0, rf)
    for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rf.set(qn(a), font)
    for tag in ("w:rtl", "w:szCs"):
        if rPr.find(qn(tag)) is None:
            el = OxmlElement(tag)
            if tag == "w:szCs":
                el.set(qn("w:val"), str(int(size_pt * 2)))
            rPr.append(el)
    if bold and rPr.find(qn("w:bCs")) is None:
        rPr.append(OxmlElement("w:bCs"))


def _style_heading(doc, name: str, font: str, size: float, color: str):
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Pt, RGBColor
    st = doc.styles[name]
    st.font.name, st.font.size, st.font.bold = font, Pt(size), True
    st.font.color.rgb = RGBColor.from_string(color)
    rPr = st.element.get_or_add_rPr()
    rf = rPr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rPr.insert(0, rf)
    for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rf.set(qn(a), font)
    for a in ("w:asciiTheme", "w:hAnsiTheme", "w:cstheme", "w:eastAsiaTheme"):
        if rf.get(qn(a)) is not None:
            del rf.attrib[qn(a)]
    st.paragraph_format.space_before = Pt(10)
    st.paragraph_format.space_after = Pt(6)
    st.paragraph_format.keep_with_next = True


def new_document(font: str, landscape: bool, margin_cm: float):
    from docx import Document
    from docx.enum.section import WD_ORIENT
    from docx.shared import Cm
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)  # A4
    if landscape:
        sec.orientation = WD_ORIENT.LANDSCAPE
        sec.page_width, sec.page_height = sec.page_height, sec.page_width
    for side in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(sec, side, Cm(margin_cm))
    normal = doc.styles["Normal"]
    normal.font.name = font
    _style_heading(doc, "Heading 1", font, 26, "1F3864")
    _style_heading(doc, "Heading 2", font, 18, "7F6000")
    _style_heading(doc, "Title", font, 30, "1F3864")
    return doc


def add_page_number_footer(doc, font: str):
    """Centered '— N —' in the footer (same as the scraper's docx files)."""
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Cm
    sec = doc.sections[0]
    sec.footer_distance = Cm(0.5)
    par = sec.footer.paragraphs[0]
    par.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_rtl_par(par)
    _font_run(par.add_run("— "), font, 14, "555555")
    r = par.add_run()
    _font_run(r, font, 14, "555555")
    for kind, txt in (("begin", None), (None, "PAGE"), ("end", None)):
        if kind:
            el = OxmlElement("w:fldChar")
            el.set(qn("w:fldCharType"), kind)
        else:
            el = OxmlElement("w:instrText")
            el.set(qn("xml:space"), "preserve")
            el.text = txt
        r._r.append(el)
    _font_run(par.add_run(" —"), font, 14, "555555")


def usable_cm(doc):
    sec = doc.sections[0]
    return (sec.page_width - sec.left_margin - sec.right_margin) / 360000, \
           (sec.page_height - sec.top_margin - sec.bottom_margin) / 360000


def add_heading(doc, text: str, level: int, font: str, size: float, color: str, align_center=True,
                page_break_before=False):
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    p = doc.add_paragraph(style=f"Heading {level}")
    _set_rtl_par(p)
    if align_center:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.page_break_before = page_break_before
    _font_run(p.add_run(text), font, size, color, bold=True)
    return p


def add_ayah_text(doc, text: str, font: str, size: float):
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    p = doc.add_paragraph()
    _set_rtl_par(p)          # RTL paragraph, no explicit jc -> aligned to the RIGHT
    p.paragraph_format.keep_with_next = True
    _font_run(p.add_run(text), font, size, "1F3864")


def add_picture(doc, png: Path, w_px: int, h_px: int, px_to_cm: float, max_w: float, max_h: float,
                alt: str):
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Cm, Pt
    w_cm, h_cm = w_px * px_to_cm, h_px * px_to_cm
    k = min(1.0, max_w / w_cm, max_h / h_cm)
    w_cm, h_cm = w_cm * k, h_cm * k
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_together = True
    pic = p.add_run().add_picture(str(png), width=Cm(w_cm), height=Cm(h_cm))
    try:  # alt text
        docPr = pic._inline.docPr
        docPr.set("descr", alt)
        docPr.set("title", alt)
    except Exception:  # noqa: BLE001
        pass


def add_cover(doc, title_extra: str, ctx: dict):
    """First-page block, same layout as the scraper's other docx files:
    basmala · book title · surah name / range · source + download date."""
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Pt

    def centered(style=None, after=4):
        p = doc.add_paragraph(style=style) if style else doc.add_paragraph()
        _set_rtl_par(p)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(after)
        return p

    p = centered(after=14)
    _font_run(p.add_run("بِسْمِ اللَّهِ الرَّحْمَنِ الرَّحِيمِ"), ctx["font"], 20, "C99200")
    p = centered("Title", after=6)
    _font_run(p.add_run(SRC_TITLE), ctx["title_font"], 36, "805830", bold=True)
    if title_extra:
        p = centered(after=6)
        _font_run(p.add_run(title_extra), ctx["font"], 20, "256080")
    p = centered(after=12)
    _font_run(p.add_run(f"المصدر: {SITE_BASE}/{SRC_SLUG} | تاريخ التحميل: {time.strftime('%Y-%m-%d')}"),
              ctx["font"], 10, "888888")


def set_page_bg(doc, hex_color):
    """Colour the whole page like the scraper's docx files (Word: Design > Page Color)."""
    if not hex_color:
        return
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    body = doc.element
    bg = OxmlElement("w:background")
    bg.set(qn("w:color"), hex_color)
    body.insert(0, bg)
    st = doc.settings.element
    if st.find(qn("w:displayBackgroundShape")) is None:
        el = OxmlElement("w:displayBackgroundShape")
        zoom = st.find(qn("w:zoom"))
        if zoom is not None:
            zoom.addnext(el)
        else:
            st.insert(0, el)


def ayah_label(a1: int, a2: int) -> str:
    return f"الآية {ar_nums(a1)}" if a1 == a2 else f"الآيات {ar_nums(a1)}–{ar_nums(a2)}"


def quran_heading(quran, abs_start, surah: int, a1: int, a2: int) -> str:
    if not quran or not abs_start:
        return ""
    try:
        parts = []
        for a in range(a1, a2 + 1):
            txt = quran[abs_start[surah - 1] + a - 1][0]
            parts.append(f"{txt} ۝{ar_nums(a)}" if a2 > a1 else txt)
        return "﴿" + " ".join(parts) + "﴾"
    except Exception:  # noqa: BLE001
        return ""


def add_surah(doc, surah: int, recs: list[dict], ctx: dict, first_in_doc: bool):
    name = ctx["name_of"](surah)
    add_heading(doc, name, 1, ctx["font"], 26, "1F3864", page_break_before=not first_in_doc)
    total = len(recs)
    for i, r in enumerate(recs, 1):
        a1, a2 = int(r["a_start"]), int(r["a_end"])
        label = ayah_label(a1, a2)
        add_heading(doc, label, 2, ctx["font"], 18, "7F6000", align_center=False)
        qt = quran_heading(ctx["quran"], ctx["abs_start"], surah, a1, a2)
        if qt:
            add_ayah_text(doc, qt, ctx["quran_font"], ctx["quran_size"])
        try:
            imgs = ctx["renderer"].render(r["data"])
        except Exception as e:  # noqa: BLE001
            log(f"    [!] render failed {name} {a1}-{a2}: {e}")
            imgs = []
        if not imgs:
            from docx.enum.text import WD_ALIGN_PARAGRAPH
            p = doc.add_paragraph("(تعذّر عرض الإعراب المرسوم لهذه الآية)")
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for k, (png, w, h) in enumerate(imgs, 1):
            add_picture(doc, png, w, h, ctx["px_to_cm"], ctx["max_w"], ctx["max_h"],
                        f"{SRC_TITLE} — {name} — {label} — {k}")
        if i % 10 == 0 or i == total:
            log(f"    {name}: {i}/{total}")


def surah_ctx_doc(ctx, landscape_font=None):
    doc = new_document(ctx["font"], ctx["landscape"], ctx["margin_cm"])
    ctx["max_w"], ctx["max_h"] = usable_cm(doc)
    set_page_bg(doc, ctx["page_bg"])
    add_page_number_footer(doc, ctx["font"])
    return doc


def build(args) -> int:
    in_path = Path(args.in_dir)
    files = find_jsonl_files(in_path)
    if not files:
        log(f"[x] no iraab-graphs surah_*.jsonl found under {in_path}")
        return 1
    records = load_records(files, parse_surahs(args.surah))
    if not records:
        log("[x] no usable records (check --surah / the data files)")
        return 1

    scraper = load_scraper(args.scraper)
    suwar = getattr(scraper, "SUWAR", None)
    abs_start = getattr(scraper, "ABS_START", None)

    def name_of(s: int) -> str:
        return f"سورة {suwar[s - 1][1]}" if suwar and 0 < s <= len(suwar) else f"سورة {s}"

    quran = None
    if not args.no_quran:
        for cand in (Path(args.quran) if args.quran else None,
                     in_path / "quran.json" if in_path.is_dir() else in_path.parent / "quran.json",
                     Path.cwd() / "tafsir_output" / "quran.json"):
            if cand and cand.exists():
                try:
                    quran = json.loads(cand.read_text(encoding="utf-8"))
                    log(f"[i] ayah text from {cand}")
                    break
                except Exception:  # noqa: BLE001
                    pass

    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    renderer = Renderer(out_dir / "_png_cache", args.scale, args.pad)

    ctx = dict(
        font=args.font, title_font=args.title_font, page_bg=args.page_bg or None, landscape=not args.portrait, margin_cm=args.margin,
        name_of=name_of, quran=quran, abs_start=abs_start,
        quran_font=args.quran_font, quran_size=args.quran_size,
        px_to_cm=args.px_cm, renderer=renderer, max_w=0, max_h=0,
    )
    n_rec = sum(len(v) for v in records.values())
    log(f"[i] {len(records)} surah(s), {n_rec} verse records")

    try:
        if not args.no_per_surah:
            for s, recs in records.items():
                doc = surah_ctx_doc(ctx)
                add_cover(doc, name_of(s), ctx)
                add_surah(doc, s, recs, ctx, first_in_doc=True)
                path = out_dir / f"{safe_name(name_of(s))}.docx"
                doc.save(path)
                log(f"[✓] {path}")
        if args.merge:
            doc = surah_ctx_doc(ctx)
            stem = surahs_stem(list(records), name_of)
            add_cover(doc, "" if len(records) == 1 else stem, ctx)
            first = True
            for s, recs in records.items():
                add_surah(doc, s, recs, ctx, first_in_doc=first)
                first = False
            path = out_dir / f"{safe_name(stem)}.docx"
            doc.save(path)
            log(f"[✓] {path}")
    finally:
        renderer.close()
    return 0


def main():
    p = argparse.ArgumentParser(description="Build Word files from iraab-graphs (الإعراب المرسوم) JSONL data")
    p.add_argument("--in_dir", default="./tafsir_output",
                   help="folder containing the raw surah_NNN.jsonl files (searched recursively) or one .jsonl file")
    p.add_argument("--out_dir", default="./iraab_graphs_docx", help="where the .docx files go")
    p.add_argument("--surah", default="all", help="'all', '4' or '1-4,112-114'")
    p.add_argument("--merge", action="store_true", help="also write ONE combined docx with all selected surahs")
    p.add_argument("--no_per_surah", action="store_true", help="with --merge: skip the per-surah files")
    p.add_argument("--scale", type=float, default=2.0, help="PNG resolution multiplier (2 = crisp, 3 = print)")
    p.add_argument("--pad", type=int, default=10, help="px of blank margin around each graph")
    p.add_argument("--px_cm", type=float, default=0.0265,
                   help="cm per SVG pixel (0.0265 ≈ website size at 96 dpi); graphs wider than the page are shrunk to fit")
    p.add_argument("--portrait", action="store_true", help="portrait A4 (default: landscape, fits wide graphs)")
    p.add_argument("--margin", type=float, default=1.2, help="page margin in cm")
    p.add_argument("--font", default="KFGQPC Uthman Taha Naskh", help="font for headings, labels and cover text")
    p.add_argument("--title_font", default="KFGQPC Uthman Taha Naskh", help="font for the book title on the cover")
    p.add_argument("--page_bg", default="", help="page colour hex, e.g. F3EEE2 (the scraper's beige); default '' = white page")
    p.add_argument("--quran_font", default="Kitab", help="font for the ﴿ayah﴾ line")
    p.add_argument("--quran_size", type=float, default=22)
    p.add_argument("--no_quran", action="store_true", help="don't print ﴿ayah﴾ text above the graphs")
    p.add_argument("--quran", default=None, help="path to quran.json (default: <in_dir>/quran.json)")
    p.add_argument("--scraper", default=None, help="path to tafsir_scraper.py (for surah names)")
    args = p.parse_args()
    if args.no_per_surah and not args.merge:
        p.error("--no_per_surah only makes sense with --merge")
    sys.exit(build(args))


if __name__ == "__main__":
    main()