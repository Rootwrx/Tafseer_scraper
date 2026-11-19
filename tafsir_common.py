"""
tafsir_common.py — pieces shared by tafsir_scraper.py (download) and tafsir_docx.py (Word output).

  • the 114-surah table, the built-in source lists and Arabic titles (sources_ar.json registry)
  • the output-folder layout (<group>/<category>/<title>/raw|docx) and the small helpers
    for saved raw data (raw_paths / load_records), surah-range parsing, logging, JSON I/O
  • reading a source's cover-page info (المؤلف / العقيدة) from the cached sources.json

Nothing in here touches the network and nothing in here builds a Word file.
"""

import html as _html
import json
import re
import threading
from html.parser import HTMLParser
from pathlib import Path

BASE = "https://tafsir.app"

# ═══════════════════════════════════════════════════════════════════════════
# STATIC DATA (copied from the site's own JS bundle / HTML)
# ═══════════════════════════════════════════════════════════════════════════

# [ayah_count, name] for surahs 1..114
SUWAR = [
    (7, "الفاتحة"), (286, "البقرة"), (200, "آل عمران"), (176, "النساء"), (120, "المائدة"),
    (165, "الأنعام"), (206, "الأعراف"), (75, "الأنفال"), (129, "التوبة"), (109, "يونس"),
    (123, "هود"), (111, "يوسف"), (43, "الرعد"), (52, "إبراهيم"), (99, "الحجر"),
    (128, "النحل"), (111, "الإسراء"), (110, "الكهف"), (98, "مريم"), (135, "طه"),
    (112, "الأنبياء"), (78, "الحج"), (118, "المؤمنون"), (64, "النور"), (77, "الفرقان"),
    (227, "الشعراء"), (93, "النمل"), (88, "القصص"), (69, "العنكبوت"), (60, "الروم"),
    (34, "لقمان"), (30, "السجدة"), (73, "الأحزاب"), (54, "سبأ"), (45, "فاطر"),
    (83, "يس"), (182, "الصافات"), (88, "ص"), (75, "الزمر"), (85, "غافر"),
    (54, "فصلت"), (53, "الشورى"), (89, "الزخرف"), (59, "الدخان"), (37, "الجاثية"),
    (35, "الأحقاف"), (38, "محمد"), (29, "الفتح"), (18, "الحجرات"), (45, "ق"),
    (60, "الذاريات"), (49, "الطور"), (62, "النجم"), (55, "القمر"), (78, "الرحمن"),
    (96, "الواقعة"), (29, "الحديد"), (22, "المجادلة"), (24, "الحشر"), (13, "الممتحنة"),
    (14, "الصف"), (11, "الجمعة"), (11, "المنافقون"), (18, "التغابن"), (12, "الطلاق"),
    (12, "التحريم"), (30, "الملك"), (52, "القلم"), (52, "الحاقة"), (44, "المعارج"),
    (28, "نوح"), (28, "الجن"), (20, "المزمل"), (56, "المدثر"), (40, "القيامة"),
    (31, "الإنسان"), (50, "المرسلات"), (40, "النبأ"), (46, "النازعات"), (42, "عبس"),
    (29, "التكوير"), (19, "الانفطار"), (36, "المطففين"), (25, "الانشقاق"), (22, "البروج"),
    (17, "الطارق"), (19, "الأعلى"), (26, "الغاشية"), (30, "الفجر"), (20, "البلد"),
    (15, "الشمس"), (21, "الليل"), (11, "الضحى"), (8, "الشرح"), (8, "التين"),
    (19, "العلق"), (5, "القدر"), (8, "البينة"), (8, "الزلزلة"), (11, "العاديات"),
    (11, "القارعة"), (8, "التكاثر"), (3, "العصر"), (9, "الهمزة"), (5, "الفيل"),
    (4, "قريش"), (7, "الماعون"), (3, "الكوثر"), (6, "الكافرون"), (3, "النصر"),
    (5, "المسد"), (4, "الإخلاص"), (5, "الفلق"), (6, "الناس"),
]
ABS_START = [0]
for _c, _n in SUWAR:
    ABS_START.append(ABS_START[-1] + _c)  # ABS_START[s-1] + a == absolute ayah id

# slug -> short Arabic title (the "تفسير" tab)
TAFSIR_SOURCES = {
    "tabari": "تفسير الطبري", "ibn-katheer": "تفسير ابن كثير", "qurtubi": "تفسير القرطبي",
    "baghawi": "تفسير البغوي", "zad-almaseer": "تفسير ابن الجوزي", "almawirdee": "تفسير الماوردي",
    "ibn-alqayyim": "تفسير ابن القيم", "ibn-taymiyyah": "تفسير ابن تيمية",
    "samaani": "تفسير السمعاني", "makki": "تفسير مكّي", "mahasin-altaweel": "محاسن التأويل للقاسمي",
    "althaalabi": "تفسير الثعالبي", "samarqandi": "تفسير السمرقندي", "althalabi": "تفسير الثعلبي",
    "fath-albayan": "فتح البيان للقنوجي", "fath-alqadeer": "فتح القدير للشوكاني",
    "altasheel": "تفسير ابن جزي", "alaloosi": "تفسير الآلوسي", "alrazi": "تفسير الرازي",
    "adwaa-albayan": "أضواء البيان", "nathm-aldurar": "نظم الدرر للبقاعي",
    "ibn-aashoor": "التحرير والتنوير", "ibn-atiyah": "المحرر الوجيز لابن عطية",
    "albahr-almuheet": "البحر المحيط لأبي حيان", "albaseet": "البسيط للواحدي",
    "abu-alsuod": "تفسير أبي السعود", "kashaf": "الكشاف للزمخشري", "muyassar": "الميسر",
    "mukhtasar": "المختصر في التفسير", "saadi": "تفسير السعدي", "aysar-altafasir": "أيسر التفاسير",
    "tadabbur-wa-amal": "القرآن – تدبر وعمل", "ibn-uthaymeen": "تفسير ابن عثيمين",
    "jalalayn": "تفسير الجلالين", "iejee": "جامع البيان للإيجي", "albaydawee": "تفسير البيضاوي",
    "alnasafi": "تفسير النسفي", "alwajeez": "الوجيز للواحدي", "zimneen": "تفسير ابن أبي زمنين",
    "mathoor": "موسوعة التفسير المأثور", "aldur-almanthoor": "الدر المنثور",
    "ibn-abi-hatim": "تفسير ابن أبي حاتم", "muqatil": "تفسير مقاتل", "qatada": "تفسير قتادة",
}

# the "علوم" tab: غريب ومعاني / أسباب النزول / إعراب ولغة / قراءات / أحكام
ULOOM_SOURCES = {
    "qiraat-almawsoah": "القراءات — الموسوعة القرآنية",
    "mafateeh-alaghanee": "مفاتيح الأغاني في القراءات",
    "ahkam-altarayfi": "أحكام القرآن للطريفي",
    "alikleel": "الإكليل للسيوطي",
    "eedah-naskh-mansukh": "ناسخ القرآن ومنسوخه لمكي",
    "siraaj-ghareeb": "غريب القرآن للخضيري", "almuyassar-ghareeb": "الميسر في الغريب",
    "ghareeb-ibn-qutaybah": "غريب القرآن لابن قتيبة",
    "altibyan-ghreeb": "غريب القرآن لابن الهائم", "zajjaj": "معاني الزجاج",
    "nahaas-meanings": "معاني القرآن للنحاس", "farraa": "معاني القرآن للفراء",
    "majaz-alquran": "مجاز القرآن لمعمر بن المثنى", "akhfash": "معاني القرآن للأخفش",
    "wahidi-asbab": "أسباب النزول للواحدي", "aliraab-almuyassar": "الإعراب الميسر",
    "iraab-daas": "إعراب القرآن للدعاس", "aljadwal": "الجدول في إعراب القرآن",
    "aldur-almasoon": "الدر المصون للسمين الحلبي", "lubab": "اللباب في علوم الكتاب",
    "iraab-aldarweesh": "إعراب القرآن للدرويش",
    "mujtaba-mushkil-iraab": "مجتبى مشكل إعراب القرآن",
    "iraab-alnahas": "إعراب القرآن للنحاس", "ayah-morph": "تحليل كلمات القرآن",
    "tahbeer-altayseer": "تحبير التيسير لابن الجزري", "alnashir": "النشر لابن الجزري",
    "ahkam-ibn-alarabee": "أحكام القرآن لابن العربي", "aljasas": "أحكام القرآن للجصاص",
    "ilkia-alharrasee": "أحكام القرآن لإلكيا الهراسي",
}
# not plain text (returns HTML/graphs) -> not part of --group uloom, but it still gets an
# Arabic title if you pass it explicitly with --src.
EXTRA_SOURCES = {"iraab-graphs": "الإعراب المرسوم"}
GROUPS = {"tafsir": TAFSIR_SOURCES, "uloom": ULOOM_SOURCES}
ALL_SOURCES = {**TAFSIR_SOURCES, **ULOOM_SOURCES, **EXTRA_SOURCES}

# ── external Arabic registry (sources_ar.json, next to this script / in the cwd) ──
# {slug: {"title": "<short Arabic title>", "group": "tafsir|uloom", "category": "<optional>", "enabled": true, ...}}
# Edit that file to add a new source: it is merged into the built-in lists above.
SOURCES_FILE = None   # set from --sources_file


def load_source_registry(extra: str | None = None) -> int:
    paths = [Path(__file__).resolve().with_name("sources_ar.json"), Path("sources_ar.json")]
    if extra:
        paths.append(Path(extra))
    n = 0
    for pth in paths:
        try:
            data = json.loads(pth.read_text(encoding="utf-8"))
        except Exception:
            continue
        if not isinstance(data, dict):
            continue
        for slug, meta in data.items():
            title = meta.get("title") if isinstance(meta, dict) else meta
            if not title:
                continue
            ALL_SOURCES[slug] = title
            if isinstance(meta, dict) and meta.get("category"):
                CATEGORY_OVERRIDE[slug] = str(meta["category"])
            grp = meta.get("group") if isinstance(meta, dict) else None
            if grp in GROUPS and meta.get("enabled", True) and slug not in EXTRA_SOURCES:
                GROUPS[grp].setdefault(slug, title)
            n += 1
    return n


def src_title(src: str) -> str:
    """Arabic display name of a source (never the English slug when it is known)."""
    return ALL_SOURCES.get(src, src)


# ═══════════════════════════════════════════════════════════════════════════
# OUTPUT FOLDERS  (source / combination folders in Arabic; raw and docx stay English)
# ═══════════════════════════════════════════════════════════════════════════

COMBINED_DIR = "جمع_التفاسير"
# technical sub-folders stay in English (raw / docx); the Arabic names
# below are only recognised so folders created by an earlier version are renamed back.
_AR_SUBDIR = {"raw": "البيانات_الخام", "docx": "وورد"}
_fs_lock = threading.Lock()


def safe_name(text: str, limit: int = 150) -> str:
    """Folder-safe version of an Arabic title (keeps Arabic, drops \\ / : * ? " < > |)."""
    t = re.sub(r'[\\/:*?"<>|\r\n\t]+', " ", text)
    t = re.sub(r"\s+", " ", t).strip(" .")
    if len(t) > limit:
        import hashlib
        t = t[:limit - 9].rstrip(" ._") + "_" + hashlib.md5(text.encode("utf-8")).hexdigest()[:8]
    return t or "مصدر"


def _adopt(old: Path, new: Path) -> Path:
    """Rename / move a legacy folder to its new place once, then return `new`."""
    with _fs_lock:
        if old != new and old.is_dir() and not new.exists():
            try:
                new.parent.mkdir(parents=True, exist_ok=True)   # <group>/<category>/ may not exist yet
                old.rename(new)
                log(f"  [folders] moved {old.name} → {new.relative_to(old.parent) if old.parent in new.parents else new.name}")
            except OSError as e:
                log(f"[!] could not rename {old} → {new} ({e}); using the old folder")
                return old
        new.parent.mkdir(parents=True, exist_ok=True)
    return new


# ─── <group>/<category>/<source> layout (same rules as transform.py) ───────────────
GROUP_DIR = {"tafsir": "تفسير", "uloom": "علوم"}     # folder names of the two site tabs
UNCATEGORIZED = "غير مصنف"
LAYOUT = "grouped"          # set from --layout: grouped | flat

# the site's own category tabs (same order as on tafsir.app); --html re-reads them from a saved page
CATEGORIES = {
    "tafsir": {
        "أمّهات": ["tabari", "ibn-katheer", "qurtubi", "baghawi"],
        "جمع الأقوال": ["zad-almaseer", "almawirdee"],
        "منتقاة": ["ibn-alqayyim", "ibn-taymiyyah"],
        "عامّة": ["samaani", "makki", "mahasin-altaweel", "althaalabi", "samarqandi",
                  "althalabi", "fath-albayan", "fath-alqadeer", "altasheel"],
        "موسوعات": ["alaloosi", "alrazi"],
        "أخرى": ["adwaa-albayan", "nathm-aldurar"],
        "لغة وبلاغة": ["ibn-aashoor", "ibn-atiyah", "albahr-almuheet", "albaseet",
                       "abu-alsuod", "kashaf"],
        "معاصرة": ["muyassar", "mukhtasar", "saadi", "aysar-altafasir",
                   "tadabbur-wa-amal", "ibn-uthaymeen"],
        "مركَّزة العبارة": ["jalalayn", "iejee", "albaydawee", "alnasafi", "alwajeez", "zimneen"],
        "آثار": ["mathoor", "aldur-almanthoor", "ibn-abi-hatim", "muqatil", "qatada"],
    },
    "uloom": {
        "غريب ومعاني": ["siraaj-ghareeb", "almuyassar-ghareeb", "ghareeb-ibn-qutaybah",
                        "altibyan-ghreeb", "zajjaj", "nahaas-meanings", "farraa",
                        "majaz-alquran", "akhfash"],
        "أسباب النزول": ["wahidi-asbab"],
        "إعراب ولغة": ["aliraab-almuyassar", "iraab-daas", "aljadwal", "aldur-almasoon",
                       "lubab", "iraab-aldarweesh", "mujtaba-mushkil-iraab",
                       "iraab-alnahas", "ayah-morph", "iraab-graphs"],
        "قراءات": ["qiraat-almawsoah", "tahbeer-altayseer", "alnashir", "mafateeh-alaghanee"],
        "أحكام": ["ahkam-ibn-alarabee", "aljasas", "ilkia-alharrasee", "ahkam-altarayfi",
                  "alikleel", "eedah-naskh-mansukh"],
    },
}
_CAT_INDEX: dict[str, dict[str, str]] = {}
CATEGORY_OVERRIDE: dict[str, str] = {}   # slug -> category, from sources_ar.json ("category" key)
SITE_META: dict = {}                     # cached sources.json (fallback for slugs not built in)


def set_categories(cats: dict | None = None) -> None:
    """(Re)build slug → category from CATEGORIES, or from a table parsed out of a saved page."""
    _CAT_INDEX.clear()
    for g, d in (cats or CATEGORIES).items():
        _CAT_INDEX[g] = {s: c for c, slugs in d.items() for s in slugs}


set_categories()


def set_site_meta(raw) -> None:
    global SITE_META
    SITE_META = raw if isinstance(raw, dict) else {}


def source_group(src: str) -> str | None:
    """'tafsir' | 'uloom' | None (None = stays at the top level, e.g. iraab-graphs)."""
    if src in EXTRA_SOURCES:
        return None
    if src in TAFSIR_SOURCES:
        return "tafsir"
    if src in ULOOM_SOURCES:
        return "uloom"
    meta = SITE_META.get(src)
    g = meta.get("group") if isinstance(meta, dict) else None
    return g if g in GROUP_DIR else None


def source_category(src: str) -> str:
    if src in CATEGORY_OVERRIDE:
        return safe_name(CATEGORY_OVERRIDE[src])
    meta = SITE_META.get(src)
    if isinstance(meta, dict):
        for k in ("category", "التصنيف"):
            if meta.get(k):
                return safe_name(str(meta[k]))
    return _CAT_INDEX.get(source_group(src), {}).get(src, UNCATEGORIZED)


def layout_path(out_dir: Path, src: str) -> Path:
    """Where a source lives:  <out>/<group>/<category>/<title>   (or <out>/<title> if flat)."""
    name = safe_name(src_title(src))
    g = source_group(src)
    if LAYOUT == "flat" or g is None:
        return out_dir / name
    return out_dir / GROUP_DIR[g] / source_category(src) / name


def _legacy_names(src: str) -> list[str]:
    """Top-level folder names an earlier version / transform may have left behind."""
    names = [safe_name(src_title(src)), src]
    meta = SITE_META.get(src)
    if isinstance(meta, dict):
        for k in ("short_name", "الاسم"):
            if meta.get(k):
                names.append(safe_name(meta[k]))
    return list(dict.fromkeys(names))


def src_dir(out_dir: Path, src: str) -> Path:
    """<out_dir>/<group>/<category>/<Arabic source title>/
    An old flat folder (<out_dir>/<title>/ or <out_dir>/<slug>/) is moved there once."""
    target = layout_path(out_dir, src)
    if not target.is_dir():
        reserved = set(GROUP_DIR.values()) | {COMBINED_DIR}
        for name in _legacy_names(src):
            old = out_dir / name
            if name not in reserved and old != target and old.is_dir():
                return _adopt(old, target)
    _adopt(target, target)      # just creates the parent folders
    return target


class _PageParser(HTMLParser):
    """Reads the tab panes of a saved tafsir.app page: <h3>category</h3> then <a data-src=slug>."""
    _PANES = {"tafsir-grp": "tafsir", "uloom-grp": "uloom"}

    def __init__(self):
        super().__init__()
        self.pane = self.cat = None
        self.in_h3 = False
        self.buf = ""
        self.cats: dict[str, dict[str, list[str]]] = {"tafsir": {}, "uloom": {}}

    def handle_starttag(self, tag, attrs):
        a = dict(attrs)
        if tag == "div" and (a.get("id") or "").endswith("-grp"):
            self.pane, self.cat = self._PANES.get(a["id"]), None
        elif tag == "h3":
            self.in_h3, self.buf = True, ""
        elif tag == "a" and self.pane and self.cat and a.get("data-src"):
            lst = self.cats[self.pane].setdefault(self.cat, [])
            if a["data-src"] not in lst:
                lst.append(a["data-src"])

    def handle_data(self, data):
        if self.in_h3:
            self.buf += data

    def handle_endtag(self, tag):
        if tag == "h3" and self.in_h3:
            self.in_h3 = False
            self.cat = " ".join(self.buf.split()) or None


def parse_categories_html(path: Path) -> dict:
    p = _PageParser()
    p.feed(Path(path).read_text(encoding="utf-8"))
    return p.cats


def sub_dir(parent: Path, key: str) -> Path:
    """parent/<raw | docx>  (renames an Arabic-named one from an older version)."""
    d = _adopt(parent / _AR_SUBDIR[key], parent / key) if key in _AR_SUBDIR else parent / key
    d.mkdir(parents=True, exist_ok=True)
    return d


def combined_dir(out_dir: Path, order: list[str]) -> Path:
    """<out_dir>/جمع_التفاسير/<title1>_<title2>_…/  — one folder per combination (and order),
    so combining different sources never overwrites an earlier result."""
    name = safe_name("_".join(safe_name(src_title(x)) for x in order))
    d = _adopt(out_dir / "combined", out_dir / COMBINED_DIR) if (out_dir / "combined").is_dir() else out_dir / COMBINED_DIR
    d = d / name
    d.mkdir(parents=True, exist_ok=True)
    return d


# ═══════════════════════════════════════════════════════════════════════════
# SMALL UTILITIES
# ═══════════════════════════════════════════════════════════════════════════

_print_lock = threading.Lock()


def log(*a, **kw):
    with _print_lock:
        print(*a, **kw, flush=True)


def ar_nums(n) -> str:
    return str(n).translate(str.maketrans("0123456789", "٠١٢٣٤٥٦٧٨٩"))


def abs_id(s: int, a: int) -> int:
    return ABS_START[s - 1] + a


def atomic_write_json(path: Path, data) -> None:
    path = Path(path)
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(path)


def load_json(path: Path, default=None):
    try:
        return json.loads(Path(path).read_text(encoding="utf-8"))
    except Exception:
        return default


def parse_surah_spec(spec: str) -> list[int]:
    """'all' | '53' | '1-3,53,112-114'  ->  sorted list of surah numbers."""
    spec = spec.strip().lower()
    if spec in ("all", "*"):
        return list(range(1, 115))
    out: set[int] = set()
    for part in spec.split(","):
        part = part.strip()
        if not part:
            continue
        try:
            if "-" in part:
                lo, hi = (int(x) for x in part.split("-", 1))
                if lo > hi:
                    raise SystemExit(f"[!] invalid surah range '{part}' (start is after end)")
                out.update(range(lo, hi + 1))
            else:
                out.add(int(part))
        except ValueError:
            raise SystemExit(f"[!] cannot read surah spec '{part}' (use e.g. 53 | 1-3,53 | all)")
    bad = [s for s in out if not 1 <= s <= 114]
    if bad:
        raise SystemExit(f"[!] invalid surah number(s): {sorted(bad)} (valid: 1-114)")
    return sorted(out)


def parse_ayah_range(spec: str | None, surah: int) -> tuple[int, int]:
    n = SUWAR[surah - 1][0]
    if not spec:
        return 1, n
    try:
        if "-" in spec:
            lo, hi = (int(x) for x in spec.split("-", 1))
        else:
            lo = hi = int(spec)
    except ValueError:
        raise SystemExit(f"[!] cannot read --ayahs '{spec}' (use e.g. 33-54 or 5)")
    if lo > hi:
        raise SystemExit(f"[!] invalid --ayahs '{spec}' (start is after end)")
    lo, hi = max(1, lo), min(n, hi)
    if lo > hi:
        raise SystemExit(f"[!] --ayahs is outside surah {surah} (it has {n} ayahs)")
    return lo, hi


# HTML tags (stripped from the cover-page info and from the book intro)
_TAG_RE = re.compile(r"<[^>]+>")


# ═══════════════════════════════════════════════════════════════════════════
# SAVED RAW DATA  (written by tafsir_scraper.py, read by tafsir_docx.py)
# ═══════════════════════════════════════════════════════════════════════════


def raw_paths(out_dir: Path, src: str, surah: int) -> tuple[Path, Path]:
    d = sub_dir(src_dir(out_dir, src), "raw")
    return d / f"surah_{surah:03d}.jsonl", d / f"surah_{surah:03d}.done"


def load_records(path: Path) -> list[dict]:
    recs = []
    if path.exists():
        with open(path, encoding="utf-8") as fh:
            for line in fh:
                line = line.strip()
                if not line:
                    continue
                try:
                    recs.append(json.loads(line))
                except json.JSONDecodeError:
                    pass  # truncated last line from an interrupted run
    return recs


def quran_heading(quran, surah: int, a_start: int, a_end: int) -> str:
    """﴿ ayah ۝n ayah ۝n ﴾ — best effort, '' if text unavailable."""
    if not quran:
        return ""
    try:
        parts = []
        for a in range(a_start, a_end + 1):
            txt = quran[abs_id(surah, a) - 1][0]
            parts.append(f"{txt} ۝{ar_nums(a)}" if a_end > a_start else txt)
        return f"﴿{' '.join(parts)}﴾"
    except Exception:
        return ""


def surah_file(s: int) -> str:
    return f"سورة {SUWAR[s - 1][1]}"


def surahs_stem(surahs: list[int]) -> str:
    """File-name stem for a merged file, built from the surahs the file REALLY contains:
         114 surahs → القرآن كاملا          one surah → سورة النجم
         contiguous → سورة النجم إلى سورة الرحمن
         otherwise  → سور 1-3،53،112-114   (never claims a range that has holes)"""
    surahs = sorted(set(surahs))
    if len(surahs) == 114:
        return "القرآن كاملا"
    if len(surahs) == 1:
        return surah_file(surahs[0])
    if surahs[-1] - surahs[0] + 1 == len(surahs):
        return f"{surah_file(surahs[0])} إلى {surah_file(surahs[-1])}"
    runs, lo, prev = [], surahs[0], surahs[0]
    for x in surahs[1:] + [None]:
        if x is None or x != prev + 1:
            runs.append(str(lo) if lo == prev else f"{lo}-{prev}")
            lo = x
        prev = x
    return safe_name("سور " + "،".join(runs), limit=100)


# ── source info for the cover page: author (المؤلف) + creed (العقيدة), read from sources.json ──
INFO_FIELDS = (("المؤلف", "المؤلف"), ("العقيدة", "العقيدة"))   # (key in sources.json, label on the cover)


def _find_field(obj, key: str, depth: int = 0):
    """First non-empty value stored under *key*, looking in nested dicts/lists too."""
    if depth > 4:
        return None
    if isinstance(obj, dict):
        v = obj.get(key)
        if v not in (None, "", [], {}):
            return v
        for sub in obj.values():
            if isinstance(sub, (dict, list)):
                r = _find_field(sub, key, depth + 1)
                if r not in (None, "", [], {}):
                    return r
    elif isinstance(obj, list):
        for sub in obj:
            if isinstance(sub, (dict, list)):
                r = _find_field(sub, key, depth + 1)
                if r not in (None, "", [], {}):
                    return r
    return None


def _flat_value(v) -> str:
    if isinstance(v, str):
        s = v
    elif isinstance(v, (list, tuple)):
        s = "، ".join(x for x in (_flat_value(i) for i in v) if x)
    elif isinstance(v, dict):
        for k in ("name", "title", "text", "value", "اسم"):
            if v.get(k):
                return _flat_value(v[k])
        s = "، ".join(x for x in (_flat_value(i) for i in v.values()) if x)
    else:
        s = "" if v is None else str(v)
    s = _html.unescape(_TAG_RE.sub(" ", s))
    return re.sub(r"\s+", " ", s).strip()


def source_info(src: str) -> list[tuple[str, str]]:
    """[('المؤلف', '...'), ('العقيدة', '...')] — only the fields the source actually has."""
    meta = SITE_META.get(src) if isinstance(SITE_META, dict) else None
    out = []
    if isinstance(meta, dict):
        for key, label in INFO_FIELDS:
            val = _flat_value(_find_field(meta, key))
            if val:
                out.append((label, val))
    return out


def _clean_intro(raw) -> str:
    if not isinstance(raw, str):
        return ""
    s = re.sub(r"<br\s*/?>", "\n", raw, flags=re.I)
    s = _html.unescape(_TAG_RE.sub("", s)).replace("\u00a0", " ")
    lines = [ln.strip() for ln in s.split("\n")]
    return re.sub(r"\n{3,}", "\n\n", "\n".join(lines)).strip()
