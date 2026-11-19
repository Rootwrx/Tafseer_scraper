#!/usr/bin/env python3
"""
maajim_scraper.py — scrape the dictionaries ("معاجم") of tafsir.app (الباحث القرآني).

How the site works
──────────────────
  One request per (dictionary, ROOT):
      GET https://tafsir.app/get_word.php?src=<slug>&w=<root>   →  {"data": "...entry..."}
  A root with no entry comes back empty.  So all you need is each dictionary's own list of roots.

Step 1 — export the root lists from the site (once)
───────────────────────────────────────────────────
  Open https://tafsir.app, press F12 → Console, paste this, and move the downloaded file to
  <out_dir>/معاجم/roots_by_source.json      (default: ./tafsir_output/معاجم/roots_by_source.json)

      const slugs = ["umdah-alhufadh","mufradat-ragheb","ishtiqaqi","quran-roots","qamoos",
                     "maqayees","lisan","basair-tamyeez","alsihah"];
      const out = {};
      slugs.forEach(s => { const v = APP_DATA.arabic_words[s];
        if (Array.isArray(v) && v.length) out[s] = {roots: v}; });
      const a = document.createElement('a');
      a.href = URL.createObjectURL(new Blob([JSON.stringify(out, null, 2)], {type:'application/json'}));
      a.download = 'roots_by_source.json'; a.click();

  Accepted shapes:  {"slug": {"roots": [...]}}   or   {"slug": [...]}

Step 2 — scrape
───────────────
  python maajim_scraper.py --list                              # dictionaries + root counts
  python maajim_scraper.py --src mufradat-ragheb --probe آدم   # test ONE root (nothing saved)
  python maajim_scraper.py --src all --md                      # every dictionary in roots_by_source.json
  python maajim_scraper.py --src lisan,qamoos --md --docx --merge
  python maajim_scraper.py --src lisan --status                # progress
  Re-running resumes: roots already asked are skipped (--force starts the book over).

Other root sources (apply the SAME list to every --src of that run)
───────────────────────────────────────────────────────────────────
  --roots_file FILE   one root per line, or a JSON list
  --roots a,b,c       roots on the command line
  --brute 3           every 3-letter combination (≈24k requests per book — last resort)

Output tree
───────────
  <out_dir>/معاجم/<عنوان المعجم>/raw/words.jsonl   one line per root {"src","w","data"}  (source of truth)
  <out_dir>/معاجم/<عنوان المعجم>/raw/tried.txt     every root already asked (ok / empty)
  <out_dir>/معاجم/<عنوان المعجم>/md|txt|json|docx/حرف <حرف>.<ext>     (--md --txt --json --docx)
  --merge → ONE file per format directly inside the dictionary folder
"""

import argparse
import itertools
import json
import re
import sys
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

try:
    from curl_cffi import requests as http
    _IMPERSONATE = True
except ImportError:
    import requests as http
    _IMPERSONATE = False

try:
    import docx as _docxlib
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Cm, Pt, RGBColor
except ImportError:
    _docxlib = None

BASE = "https://tafsir.app"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (X11; Linux x86_64; rv:128.0) Gecko/20100101 Firefox/128.0",
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "ar,en-US;q=0.5",
}
GROUP_DIR = "معاجم"

# Arabic folder names (override / add with --src_title slug=العنوان)
MAAJIM = {
    "umdah-alhufadh": "عمدة الحفاظ",
    "mufradat-ragheb": "مفردات ألفاظ القرآن",
    "ishtiqaqi": "المعجم الاشتقاقي",
    "quran-roots": "فهرس جذور كلمات القرآن",
    "qamoos": "القاموس المحيط",
    "maqayees": "مقاييس اللغة",
    "lisan": "لسان العرب",
    "basair-tamyeez": "بصائر ذوي التمييز",
    "alsihah": "الصحاح",
}
ROOTS_FILE = "roots_by_source.json"
LETTERS = "أبتثجحخدذرزسشصضطظعغفقكلمنهويء"

_print_lock = threading.Lock()
_file_lock = threading.Lock()
_net_lock = threading.Lock()
_net_requests = 0


def log(*a):
    with _print_lock:
        print(*a, flush=True)


def safe_name(text: str) -> str:
    t = re.sub(r'[\\/:*?"<>|\r\n\t]+', " ", text)
    return re.sub(r"\s+", " ", t).strip(" .") or "معجم"


def ar_nums(n) -> str:
    return str(n).translate(str.maketrans("0123456789", "٠١٢٣٤٥٦٧٨٩"))


_DIAC = re.compile("[\u064B-\u0652\u0670\u0640]")


def norm_root(w: str) -> str:
    return _DIAC.sub("", w).strip()


def load_json(path: Path, default=None):
    try:
        return json.loads(Path(path).read_text(encoding="utf-8"))
    except Exception:
        return default


def atomic_write_json(path: Path, data) -> None:
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(path)



# ─── HTTP ───────────────────────────────────────────────────────────────────

def make_session(ua=None):
    s = http.Session()
    s.headers.update(HEADERS)
    if ua:
        s.headers["User-Agent"] = ua
    if _IMPERSONATE:
        s.impersonate = "firefox147" if "firefox" in s.headers["User-Agent"].lower() else "chrome124"
    return s


def request(session, url, params=None, referer=None, retries=5, timeout=25):
    """GET with retry/backoff; returns the response (404 is returned, not raised)."""
    global _net_requests
    last = "unknown error"
    headers = {"Referer": referer} if referer else {}
    for attempt in range(1, retries + 1):
        try:
            with _net_lock:
                _net_requests += 1
            r = session.get(url, params=params, headers=headers, timeout=timeout)
            if r.status_code == 429 or r.status_code >= 500:
                last = f"HTTP {r.status_code}"
                time.sleep(min(30, 2 ** attempt))
                continue
            if r.status_code == 403:
                raise RuntimeError("HTTP 403 — blocked (lower --workers / raise --delay)")
            return r
        except RuntimeError:
            raise
        except Exception as e:
            last = str(e)
            time.sleep(min(30, 2 ** attempt))
    raise RuntimeError(f"{url} failed after {retries} tries: {last}")


def fetch_word(session, base, src, root):
    """→ ('ok', text) | ('empty', '').  Anything unexpected raises (and is NOT recorded)."""
    r = request(session, f"{base}/get_word.php", {"src": src, "w": root},
                referer=f"{base}/")
    if r.status_code == 404:
        return "empty", ""
    r.raise_for_status()
    try:
        js = r.json()
    except Exception:
        raise RuntimeError(f"{src}/{root}: reply is not JSON ({r.text[:80]!r}) — not saved")
    if not isinstance(js, dict) or "data" not in js:
        raise RuntimeError(f"{src}/{root}: unexpected reply {str(js)[:100]!r} — not saved")
    data = js["data"] if isinstance(js["data"], str) else ""
    return ("ok", data) if data.strip() else ("empty", "")


# ─── text cleaning (same rules as tafsir_scraper.py) ────────────────────────

_TAG_RE = re.compile(r"<[^>]+>")
_PAGE_RE = re.compile(r"\(p-[\d٠-٩]+\)")
_FN_RE = re.compile(r" ?\[\[([\s\S]*?\]*)\]\]")


_LIG = {"\uFD40": "رحمه الله", "\uFD41": "رضي الله عنه", "\uFD42": "رضي الله عنها",
        "\uFD43": "رضي الله عنهم", "\uFD44": "رضي الله عنهما", "\uFD45": "رضي الله عنهن",
        "\uFD46": "صلى الله عليه وآله", "\uFD47": "عليه السلام", "\uFD48": "عليهم السلام",
        "\uFD49": "عليهما السلام", "\uFD4A": "عليه الصلاة والسلام", "\uFD4B": "قدس سره",
        "\uFD4C": "صلى الله عليه وعلى آله وسلم", "\uFD4D": "عليها السلام",
        "\uFD4E": "تبارك وتعالى", "\uFD4F": "رحمهم الله", "\uFDFB": "جل جلاله",
        "\uFDFE": "سبحانه وتعالى", "\uFDFF": "عز وجل"}   # ﷺ (U+FDFA) is kept: most fonts have it


def clean_text(raw: str):
    if not raw:
        return "", []
    raw = "".join(_LIG.get(ch, ch) for ch in raw)
    t = re.sub(r"<br\s*/?>", "\n", raw)
    t = _TAG_RE.sub("", t)
    t = _PAGE_RE.sub("", t)
    notes = []

    def _fn(m):
        notes.append(m.group(1).replace("<br>", "\n").strip())
        return f" ({ar_nums(len(notes))})"

    t = _FN_RE.sub(_fn, t)
    t = re.sub(r"[ \t]+\n", "\n", t)
    t = re.sub(r"\n{3,}", "\n\n", t).strip()
    return t, notes



# ─── folders ────────────────────────────────────────────────────────────────

def title_of(src, titles):
    return titles.get(src) or MAAJIM.get(src) or src


def book_dir(out_dir: Path, src, titles) -> Path:
    d = out_dir / GROUP_DIR / safe_name(title_of(src, titles))
    d.mkdir(parents=True, exist_ok=True)
    return d


def raw_files(out_dir, src, titles):
    d = book_dir(out_dir, src, titles) / "raw"
    d.mkdir(exist_ok=True)
    return d / "words.jsonl", d / "tried.txt"


def load_entries(path: Path) -> dict:
    ents = {}
    if path.exists():
        for line in path.read_text(encoding="utf-8").splitlines():
            try:
                r = json.loads(line)
                ents[r["w"]] = r
            except Exception:
                pass          # truncated last line of an interrupted run
    return ents


def load_tried(path: Path) -> set:
    if not path.exists():
        return set()
    return {l.split("\t")[0] for l in path.read_text(encoding="utf-8").splitlines() if l.strip()}



# ─── list of roots ──────────────────────────────────────────────────────────

def _roots_from_obj(obj) -> list:
    """Roots out of a list / {"roots": [...]} / dict-of-roots / list of [root, …] or {"w": root}."""
    if isinstance(obj, dict):
        for k in ("roots", "words", "keys", "data", "list"):
            if isinstance(obj.get(k), (list, dict)):
                return _roots_from_obj(obj[k])
        obj = list(obj)
    out = []
    for x in obj if isinstance(obj, list) else []:
        if isinstance(x, str):
            out.append(x)
        elif isinstance(x, (list, tuple)) and x and isinstance(x[0], str):
            out.append(x[0])
        elif isinstance(x, dict):
            out += [x[k] for k in ("w", "root", "word", "name", "key") if isinstance(x.get(k), str)][:1]
    return out


def clean_roots(lst) -> list:
    """Strip diacritics, drop empties and duplicates, keep order."""
    return list(dict.fromkeys(r for r in (norm_root(x) for x in lst) if r))


def read_roots_file(path: str) -> list:
    txt = Path(path).read_text(encoding="utf-8")
    try:
        return _roots_from_obj(json.loads(txt))
    except ValueError:
        return [l.strip() for l in txt.splitlines() if l.strip()]


def load_roots_by_source(out_dir: Path) -> dict:
    """{slug: [roots]} from <out_dir>/معاجم/roots_by_source.json (written by the console snippet)."""
    data = load_json(out_dir / GROUP_DIR / ROOTS_FILE, {})
    if not isinstance(data, dict):
        return {}
    return {k: clean_roots(_roots_from_obj(v)) for k, v in data.items() if _roots_from_obj(v)}


def build_universe(args, out_dir: Path, sources: list) -> dict:
    """{slug: [roots]}.  Explicit --roots_file/--roots/--brute win and apply to every dictionary;
    otherwise each dictionary uses its own list from roots_by_source.json."""
    if args.roots_file or args.roots or args.brute:
        roots = []
        if args.roots_file:
            roots += read_roots_file(args.roots_file)
        if args.roots:
            roots += args.roots.split(",")
        if args.brute:
            roots += ["".join(c) for c in itertools.product(args.letters, repeat=args.brute)]
        uni = clean_roots(roots)
        log(f"[*] {len(uni)} roots (same list for every dictionary)")
        return {x: uni for x in sources}
    per = load_roots_by_source(out_dir)
    missing = [x for x in sources if x not in per]
    if missing:
        log(f"[!] no root list for: {', '.join(missing)} — skipped  "
            f"(see Step 1 in the header, or use --roots_file / --brute 3)")
    return {x: per[x] for x in sources if x in per}



# ─── download ───────────────────────────────────────────────────────────────

def scrape_book(session, args, out_dir, src, titles, universe: list) -> dict:
    words_f, tried_f = raw_files(out_dir, src, titles)
    if args.force:
        words_f.write_text("", encoding="utf-8")
        tried_f.write_text("", encoding="utf-8")
    tried = load_tried(tried_f)
    todo = [r for r in universe if r not in tried]
    if not todo:
        return {"asked": 0, "found": 0, "failed": 0}
    log(f"[{src}] {len(universe) - len(todo)} already asked, {len(todo)} to ask")

    found = asked = failed = 0
    stop = threading.Event()

    def work(root):
        if stop.is_set():
            return None
        st, data = fetch_word(session, args.base, src, root)
        if args.delay:
            time.sleep(args.delay)
        return st, data

    pool = ThreadPoolExecutor(max_workers=args.workers)
    try:
        with open(words_f, "a", encoding="utf-8") as wfh, open(tried_f, "a", encoding="utf-8") as tfh:
            futs = {pool.submit(work, r): r for r in todo}
            for i, f in enumerate(as_completed(futs), 1):
                root = futs[f]
                try:
                    res = f.result()
                except Exception as e:
                    failed += 1
                    if failed <= 10:
                        log(f"  [{src}] ERROR {root}: {e}")
                    if "403" in str(e):
                        stop.set()
                    continue
                if res is None:
                    continue
                st, data = res
                asked += 1
                with _file_lock:
                    if st == "ok":
                        wfh.write(json.dumps({"src": src, "w": root, "data": data}, ensure_ascii=False) + "\n")
                        wfh.flush()
                        found += 1
                    tfh.write(f"{root}\t{st}\n")
                    tfh.flush()
                if i % 500 == 0:
                    log(f"  [{src}] {i}/{len(todo)} asked, {found} with an entry")
    except KeyboardInterrupt:
        log("\n[!] interrupted — progress saved, re-run to resume")
        stop.set()
        pool.shutdown(wait=False, cancel_futures=True)
        sys.exit(1)
    pool.shutdown(wait=True)
    return {"asked": asked, "found": found, "failed": failed}



# ─── output ─────────────────────────────────────────────────────────────────

def letter_of(root: str) -> str:
    c = root[:1]
    return "أ" if c in "أإآءا" else c


def _sorted_entries(ents: dict):
    order = {ch: i for i, ch in enumerate("أبتثجحخدذرزسشصضطظعغفقكلمنهوي")}
    return sorted(ents.values(), key=lambda r: ([order.get(ch, 99) for ch in r["w"]], r["w"]))


def _md_entry(r) -> str:
    body, notes = clean_text(r["data"])
    parts = [f"## {r['w']}", "", "\n\n".join(l.strip() for l in body.split("\n") if l.strip())]
    if notes:
        parts += ["", "---"] + [f"> ({ar_nums(i)}) {' '.join(n.split())}  " for i, n in enumerate(notes, 1)]
    return "\n".join(parts) + "\n"


def _txt_entry(r) -> str:
    body, notes = clean_text(r["data"])
    return "\n".join([f"[{r['w']}]", body] + [f"({ar_nums(i)}) {n}" for i, n in enumerate(notes, 1)]) + "\n"


def _json_entry(r) -> dict:
    body, notes = clean_text(r["data"])
    return {"w": r["w"], "text": body, "footnotes": notes, "data": r["data"]}


# ─── DOCX (python-docx: RTL, real Heading styles → Navigation Pane, footnotes, page numbers) ───
DOCX_COLORS = {"title": "805830", "letter": "256080", "root": "B8451A", "ink": "1A1008", "note": "555555"}
VERSE_COLOR = "0F6B9C"
POEM_COLOR = "C00000"
POEM_MARK = "؎"
LINE_SPACING = 0.9
PAGE_BG = "F3EEE2"
TEXT_FONT = "KFGQPC Uthman Taha Naskh"

_PPR_AFTER_BIDI = ("w:adjustRightInd", "w:snapToGrid", "w:spacing", "w:ind", "w:contextualSpacing",
                   "w:mirrorIndents", "w:suppressOverlap", "w:jc", "w:textDirection", "w:textAlignment",
                   "w:textboxTightWrap", "w:outlineLvl", "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr",
                   "w:pPrChange")
_RPR_AFTER_BCS = ("w:i", "w:iCs", "w:caps", "w:smallCaps", "w:strike", "w:dstrike", "w:outline",
                  "w:shadow", "w:emboss", "w:imprint", "w:noProof", "w:snapToGrid", "w:vanish",
                  "w:webHidden", "w:color", "w:spacing", "w:w", "w:kern", "w:position", "w:sz",
                  "w:szCs", "w:highlight", "w:u", "w:effect", "w:bdr", "w:shd", "w:fitText",
                  "w:vertAlign", "w:rtl", "w:cs", "w:em", "w:lang", "w:eastAsianLayout",
                  "w:specVanish", "w:oMath")
_RPR_AFTER_SZCS = _RPR_AFTER_BCS[_RPR_AFTER_BCS.index("w:szCs") + 1:]
_RPR_AFTER_ICS = _RPR_AFTER_BCS[_RPR_AFTER_BCS.index("w:iCs") + 1:]
_VERSE_RE = re.compile("([﴾﴿][^﴾﴿]*[﴾﴿])")
_NUM_RE = re.compile(r"(\(\s*[\d٠-٩]+\s*\))")


def _d_bidi(par) -> None:
    pf = par.paragraph_format
    pf.left_indent = pf.right_indent = pf.first_line_indent = Pt(0)
    ppr = (par._p if hasattr(par, "_p") else par.element).get_or_add_pPr()
    if ppr.find(qn("w:bidi")) is None:
        ppr.insert_element_before(OxmlElement("w:bidi"), *_PPR_AFTER_BIDI)


def _d_run(par, text, font, size, bold=False, color=None):
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
    szcs = OxmlElement("w:szCs")
    szcs.set(qn("w:val"), str(int(round(size * 2))))
    rpr.insert_element_before(szcs, *_RPR_AFTER_SZCS)
    bcs = OxmlElement("w:bCs")
    if not bold:
        bcs.set(qn("w:val"), "0")
    rpr.insert_element_before(bcs, *_RPR_AFTER_BCS)
    ics = OxmlElement("w:iCs")
    ics.set(qn("w:val"), "0")
    rpr.insert_element_before(ics, *_RPR_AFTER_ICS)
    if color:
        r.font.color.rgb = RGBColor.from_string(color)
    return r


def _d_par(doc, text, font, size, *, bold=False, color=None, align=None, before=0, after=4):
    par = doc.add_paragraph()
    _d_bidi(par)
    par.alignment = align if align is not None else WD_ALIGN_PARAGRAPH.LEFT   # LEFT == "start" in RTL
    par.paragraph_format.space_before = Pt(before)
    par.paragraph_format.space_after = Pt(after)
    if text:
        _d_run(par, text, font, size, bold, color)
    return par


def _d_heading(doc, text, level, font, size, color, page_break=False, bold=True):
    h = doc.add_heading(level=level)
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


def _d_footer(doc, font):
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


def _d_setup(doc, font, size):
    if PAGE_BG:
        body = doc.element
        for old in body.findall(qn("w:background")):
            body.remove(old)
        bg = OxmlElement("w:background")
        bg.set(qn("w:color"), PAGE_BG)
        body.insert(0, bg)
        st = doc.settings.element
        if st.find(qn("w:displayBackgroundShape")) is None:
            el = OxmlElement("w:displayBackgroundShape")
            zoom = st.find(qn("w:zoom"))
            zoom.addnext(el) if zoom is not None else st.insert(0, el)
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.top_margin, sec.bottom_margin = Cm(1.8), Cm(1.3)
    sec.footer_distance = Cm(0.5)
    sec.left_margin = sec.right_margin = Cm(1.8)
    sp = sec._sectPr
    if sp.find(qn("w:bidi")) is None:
        sp.insert_element_before(OxmlElement("w:bidi"), "w:rtlGutter", "w:docGrid",
                                 "w:printerSettings", "w:sectPrChange")
    normal = doc.styles["Normal"]
    normal.font.name = font
    normal.font.size = Pt(size)
    normal.paragraph_format.line_spacing = LINE_SPACING
    normal.element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:cs"), font)
    _d_bidi(normal)
    for name in ("Title", "Heading 1", "Heading 2", "Heading 3"):
        hs = doc.styles[name]
        hs.font.italic = False
        hs.paragraph_format.line_spacing = LINE_SPACING
        hs.element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:cs"), font)
        hs.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        _d_bidi(hs)
    _d_footer(doc, font)


def _d_text_runs(par, text, font, verse_font, sz, color):
    """Quran quotations ﴿…﴾ in the verse font/colour, (١)-style numbers in the verse colour."""
    for i, seg in enumerate(_VERSE_RE.split(text)):
        if not seg:
            continue
        if i % 2:
            _d_run(par, seg, verse_font, sz - 2, color=VERSE_COLOR)
            continue
        for part in _NUM_RE.split(seg):
            if part:
                _d_run(par, part, font, sz, color=VERSE_COLOR if _NUM_RE.fullmatch(part) else color)


def build_docx(out_path: Path, title: str, src: str, sections: list, font: str, size: float,
               title_extra: str = "") -> None:
    """sections = [(letter, [entry, …])]"""
    if _docxlib is None:
        log("  [!] python-docx not installed — DOCX skipped (pip install python-docx)")
        return
    C = DOCX_COLORS
    doc = _docxlib.Document()
    _d_setup(doc, font, size)
    _d_par(doc, "بِسْمِ اللَّهِ الرَّحْمَنِ الرَّحِيمِ", font, size + 2, color="C99200",
           align=WD_ALIGN_PARAGRAPH.CENTER, after=14)
    t = doc.add_heading(level=0)
    for r in list(t.runs):
        r.text = ""
    _d_bidi(t)
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _d_run(t, title, TEXT_FONT, size + 14, True, C["title"])
    if title_extra:
        _d_par(doc, title_extra, font, size + 2, color=C["letter"], align=WD_ALIGN_PARAGRAPH.CENTER)
    _d_par(doc, f"المصدر: {BASE}/{src} | تاريخ التحميل: {time.strftime('%Y-%m-%d')}", font, 10,
           color="888888", align=WD_ALIGN_PARAGRAPH.CENTER, after=12)

    n = 0
    for k, (letter, entries) in enumerate(sections):
        _d_heading(doc, f"حرف {letter}", 1, font, size + 6, C["letter"], page_break=k > 0)
        for r in entries:
            n += 1
            _d_heading(doc, r["w"], 2, font, size + 2, C["root"], bold=False)
            body, notes = clean_text(r["data"])
            for line in body.split("\n"):
                line = line.strip()
                if not line:
                    continue
                if line.startswith(POEM_MARK):
                    _d_par(doc, line.lstrip(POEM_MARK).strip(), TEXT_FONT, size, color=POEM_COLOR,
                           align=WD_ALIGN_PARAGRAPH.CENTER, before=4, after=4)
                else:
                    par = _d_par(doc, "", TEXT_FONT, size, align=WD_ALIGN_PARAGRAPH.JUSTIFY)
                    _d_text_runs(par, line, TEXT_FONT, font, size, C["ink"])
            for i, note in enumerate(notes, 1):
                par = _d_par(doc, "", TEXT_FONT, max(10, size - 4), after=2)
                _d_run(par, f"({ar_nums(i)}) ", TEXT_FONT, max(10, size - 4), color=VERSE_COLOR)
                _d_text_runs(par, " ".join(note.split()), TEXT_FONT, font, max(10, size - 4), C["note"])
    doc.save(str(out_path))
    log(f"  [docx] saved → {out_path}  ({len(sections)} letter(s), {n} roots)")


def build_outputs(args, out_dir, src, titles) -> None:
    ents = load_entries(raw_files(out_dir, src, titles)[0])
    if not ents:
        log(f"  [!] {title_of(src, titles)}: no entries saved — nothing built")
        return
    title = title_of(src, titles)
    base = book_dir(out_dir, src, titles)
    ordered = _sorted_entries(ents)
    by_letter: dict = {}
    for r in ordered:
        by_letter.setdefault(letter_of(r["w"]), []).append(r)

    def write(fmt, ext, name, text):
        d = base if args.merge else base / fmt
        d.mkdir(exist_ok=True)
        (d / f"{name}.{ext}").write_text(text, encoding="utf-8")

    for fmt, ext, enc, head in (("md", "md", _md_entry, f"# {title}\n\n> المصدر: {BASE}/{src}\n\n"),
                                ("txt", "txt", _txt_entry, "")):
        if not getattr(args, fmt):
            continue
        if args.merge:
            write(fmt, ext, title, head + "\n".join(enc(r) for r in ordered))
        else:
            for ch, rs in by_letter.items():
                write(fmt, ext, f"حرف {ch}", "\n".join(enc(r) for r in rs))
    if args.json:
        meta = {"source": src, "title": title, "origin": f"{BASE}/{src}"}
        if args.merge:
            atomic_write_json(base / f"{title}.json", {**meta, "entries": [_json_entry(r) for r in ordered]})
        else:
            (base / "json").mkdir(exist_ok=True)
            for ch, rs in by_letter.items():
                atomic_write_json(base / "json" / f"حرف {ch}.json",
                                  {**meta, "entries": [_json_entry(r) for r in rs]})
    if args.docx:
        if args.merge:
            build_docx(base / f"{title}.docx", title, src, list(by_letter.items()), args.font,
                       args.docx_size)
        else:
            (base / "docx").mkdir(exist_ok=True)
            for ch, rs in by_letter.items():
                build_docx(base / "docx" / f"حرف {ch}.docx", title, src, [(ch, rs)], args.font,
                           args.docx_size, title_extra=f"حرف {ch}")
    log(f"  [{title}] {len(ents)} root entries built → {base}")



# ─── CLI ────────────────────────────────────────────────────────────────────

def main():
    p = argparse.ArgumentParser(description="Scrape tafsir.app dictionaries (معاجم)")
    p.add_argument("--src", help="comma-separated dictionary slugs, or 'all' (= every slug in roots_by_source.json)")
    p.add_argument("--src_title", action="append", default=[], metavar="SLUG=العنوان",
                   help="Arabic folder name of a dictionary (repeatable)")
    p.add_argument("--list", action="store_true", help="show dictionaries and their root counts")
    p.add_argument("--status", action="store_true", help="show download progress")
    p.add_argument("--probe", metavar="ROOT", help="fetch ONE root and print it (nothing saved)")
    g = p.add_argument_group("roots (default: <out_dir>/معاجم/roots_by_source.json)")
    g.add_argument("--roots_file", help="one root per line, or a JSON list")
    g.add_argument("--roots", help="comma-separated roots")
    g.add_argument("--brute", type=int, choices=[2, 3], default=0, help="every N-letter combination (slow)")
    g.add_argument("--letters", default=LETTERS)
    g = p.add_argument_group("download")
    g.add_argument("--out_dir", default="./tafsir_output")
    g.add_argument("--workers", type=int, default=4)
    g.add_argument("--delay", type=float, default=0.2)
    g.add_argument("--force", action="store_true", help="forget what was saved for the book and start over")
    g.add_argument("--ua", default=None)
    g.add_argument("--base", default=BASE, help=argparse.SUPPRESS)
    g = p.add_argument_group("output")
    g.add_argument("--md", action="store_true")
    g.add_argument("--txt", action="store_true")
    g.add_argument("--json", action="store_true")
    g.add_argument("--docx", action="store_true", help="Word files (needs python-docx)")
    g.add_argument("--merge", action="store_true", help="ONE file per format instead of one per first letter")
    g.add_argument("--font", default="Kitab", help="DOCX font for headings / verses")
    g.add_argument("--tafsir_font", default="KFGQPC Uthman Taha Naskh", help="DOCX font of the entry text")
    g.add_argument("--docx_size", type=float, default=16, help="DOCX body size in pt")
    g.add_argument("--line_spacing", type=float, default=0.9)
    g.add_argument("--page_bg", default="F3EEE2", help="hex page colour, or 'none' for white")
    args = p.parse_args()

    titles = dict(MAAJIM)
    for kv in args.src_title:
        if "=" in kv:
            k, v = kv.split("=", 1)
            titles[k.strip()] = v.strip()

    out_dir = Path(args.out_dir)
    per = load_roots_by_source(out_dir)

    if args.list:
        if not per:
            raise SystemExit(f"[!] {out_dir / GROUP_DIR / ROOTS_FILE} not found or empty — see Step 1 in the header")
        for slug, roots in per.items():
            print(f"  {slug:20s} {len(roots):6d} roots   {title_of(slug, titles)}")
        return

    sources = list(dict.fromkeys(s.strip() for s in (args.src or "").split(",") if s.strip()))
    if "all" in sources:
        sources = [s for s in sources if s != "all"] + [s for s in per if s not in sources]
    if not sources:
        p.error("need --src slug1,slug2 or --src all  (see --list)")

    if args.probe:
        w = norm_root(args.probe)
        st, data = fetch_word(make_session(args.ua), args.base, sources[0], w)
        print(f"[{sources[0]} / {w}] {st}")
        if data:
            body, notes = clean_text(data)
            print(body[:3000])
            print(f"... ({len(body)} chars, {len(notes)} footnotes)")
        return

    if args.status:
        for src in sources:
            wf, tf = raw_files(out_dir, src, titles)
            print(f"  {src:20s} entries {len(load_entries(wf)):6d}   asked {len(load_tried(tf)):6d}"
                  f" / {len(per.get(src, [])):6d}")
        return

    global LINE_SPACING, PAGE_BG, TEXT_FONT
    LINE_SPACING = args.line_spacing
    PAGE_BG = None if args.page_bg.lower() in ("none", "off", "white", "") else args.page_bg.lstrip("#").upper()
    TEXT_FONT = args.tafsir_font.split(",")[0].strip() or TEXT_FONT
    args.font = args.font.split(",")[0].strip() or "Kitab"
    if args.docx and _docxlib is None:
        log("[!] python-docx not installed — DOCX output disabled (pip install python-docx)")
        args.docx = False

    universes = build_universe(args, out_dir, sources)
    sources = [x for x in sources if x in universes]
    if not sources:
        raise SystemExit("[!] nothing to do")

    session = make_session(args.ua)
    for src in sources:
        res = scrape_book(session, args, out_dir, src, titles, universes[src])
        log(f"[{src}] asked {res['asked']}, with entry {res['found']}"
            + (f", FAILED {res['failed']} (re-run to retry)" if res["failed"] else ""))
        if res["asked"] and res["found"] == 0:
            log(f"  [!] {src}: every answer was empty — wrong slug or root spelling (try --probe)")

    if args.md or args.txt or args.json or args.docx:
        log("[*] building outputs ...")
        for src in sources:
            build_outputs(args, out_dir, src, titles)
    else:
        log("[*] raw data saved — add --md / --docx / --txt / --json to build files")
    log(f"[✓] network requests: {_net_requests} → {out_dir / GROUP_DIR}")


if __name__ == "__main__":
    main()