#!/usr/bin/env python3
"""
tafsir_scraper.py — Scrape tafsir / ulum sources from tafsir.app (الباحث القرآني)
and export them to JSONL + JSON + Markdown (+ plain TXT).

How the site works (from the captured traffic)
──────────────────────────────────────────────
  • Page URL     : https://tafsir.app/<src>/<surah>/<ayah>      (popup opens)
  • Data request : GET https://tafsir.app/get.php?src=<src>&s=<surah>&a=<ayah>&ver=1
                   → {"data": "...text..."}  (+ "ayahs_start" and "count" when the
                     source comments on a GROUP of ayahs in one block)
  • Metadata     : /sources/sources.json   (per-source "keys" / "keys_flag" =
                     which ayahs actually have material)
  • Quran text   : /sources/quran.json     (optional, for ﴿…﴾ headings)

Design (mirrors the Shamela scraper)
────────────────────────────────────
  • Resumable   — one append-only .jsonl per (source, surah) is the source of
                  truth; a .done marker is written when a surah finishes.
  • Offline-first — before ANY request the script looks at raw/*.jsonl, *.done and
                  the cached sources.json and works out which ayahs are really
                  missing.  Only those are requested; if nothing is missing there
                  is no network traffic at all and the files are simply rebuilt.
                  (the run ends with the number of network requests actually made)
  • Parallel    — one task per (source, surah), run on a thread pool.
  • Group-aware — if a response covers ayahs 33-54 the scraper jumps to 55
                  instead of re-requesting 34…54.
  • Skips ayahs a source has no data for (keys_flag), unless --no_prefilter.

Output tree
───────────
    <out_dir>/<عنوان المصدر>/raw/surah_053.jsonl     raw records (source of truth)
    <out_dir>/<عنوان المصدر>/raw/surah_053.done      completion marker
    <out_dir>/<عنوان المصدر>/md/سورة النجم.md        per-surah Markdown   (no format flag → md + docx;
    <out_dir>/<عنوان المصدر>/docx/سورة النجم.docx    per-surah Word        any of --md/--docx/--json/--txt → ONLY those)
    <out_dir>/<عنوان المصدر>/json/سورة النجم.json    with --json
    <out_dir>/<عنوان المصدر>/txt/سورة النجم.txt      with --txt
    --merge → ONE file per format directly inside <out_dir>/<عنوان المصدر>/ :
        سورة الفاتحة إلى سورة الناس.docx | القرآن كاملا.docx | سورة النجم إلى سورة الرحمن.md
        سور 1-3،53،112-114.docx          (non-contiguous selections are named by number)
    <out_dir>/جمع_التفاسير/<مصدر١>_<مصدر٢>/docx/سورة النجم.docx   with --combine
    (old <slug>/ folders are renamed to the Arabic source title automatically)
    <out_dir>/quran.json                    (cached; used for ﴿ayah﴾ headings)

Usage
─────
    # one ayah range of one surah, one source
    python tafsir_scraper.py --src saadi --surah 53 --ayahs 33-54

    # whole surah, several sources
    python tafsir_scraper.py --src saadi,aysar-altafasir,muyassar --surah 53

    # several surahs / ranges
    python tafsir_scraper.py --src saadi --surah 1-3,53,112-114

    # entire Quran, one source, resumable
    python tafsir_scraper.py --src saadi --surah all

    # every source of the tafsir tab for surah 53
    python tafsir_scraper.py --group tafsir --surah 53

    # Word output (RTL, real Heading styles → Navigation Pane)  — on by default
    python tafsir_scraper.py --src saadi --surah 53
    python tafsir_scraper.py --src saadi --surah 1-3 --font "Kitab" --docx_size 18

    # ONE merged file (docx + md) instead of one file per surah
    python tafsir_scraper.py --src ibn-katheer --surah all --merge
    python tafsir_scraper.py --src ibn-katheer --surah 1-9 --merge --no-md
    #   already downloaded?  run the SAME command again: it rebuilds from the saved raw/
    #   data and only requests what is missing (nothing missing → zero network requests)

    # combine several tafsirs into ONE md/docx (order = order given)
    #   per ayah: ﴿verse﴾ → saadi → ibn-alqayyim → …
    python tafsir_scraper.py --combine saadi,ibn-alqayyim,muyassar --surah 53
    python tafsir_scraper.py --combine saadi,muyassar --surah 1-3 --combine_only
    python tafsir_scraper.py --combine saadi,muyassar --surah all --merge --combine_only
    #   → <out_dir>/جمع_التفاسير/<مصدر١>_<مصدر٢>/md|docx/سورة النجم.md|.docx

    # extras
    python tafsir_scraper.py --src saadi --surah 53 --json --txt
    python tafsir_scraper.py --list_sources
    python tafsir_scraper.py --src saadi --surah 53 --force      # re-download even if saved
    python tafsir_scraper.py --src saadi --surah all --status    # progress only (no network)

Install:  pip install curl_cffi python-docx   (curl_cffi optional; python-docx only for --docx)
"""

import argparse
import html as _html
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
except ImportError:  # plain requests fallback
    import requests as http

    _IMPERSONATE = False

try:
    import docx as _docxlib
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
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
# {slug: {"title": "<short Arabic title>", "group": "tafsir|uloom", "enabled": true, ...}}
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
            grp = meta.get("group") if isinstance(meta, dict) else None
            if grp in GROUPS and meta.get("enabled", True) and slug not in EXTRA_SOURCES:
                GROUPS[grp].setdefault(slug, title)
            n += 1
    return n


def src_title(src: str) -> str:
    """Arabic display name of a source (never the English slug when it is known)."""
    return ALL_SOURCES.get(src, src)


# ═══════════════════════════════════════════════════════════════════════════
# OUTPUT FOLDERS  (source / combination folders in Arabic; raw, md, docx, txt, json stay English)
# ═══════════════════════════════════════════════════════════════════════════

COMBINED_DIR = "جمع_التفاسير"
# technical sub-folders stay in English (raw / md / docx / txt / json); the Arabic names
# below are only recognised so folders created by an earlier version are renamed back.
_AR_SUBDIR = {"raw": "البيانات_الخام", "md": "ماركداون", "docx": "وورد",
              "txt": "نصوص", "json": "جيسون"}
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
    """Rename a legacy (English-named) folder to its Arabic name once, then return `new`."""
    with _fs_lock:
        if old != new and old.is_dir() and not new.exists():
            try:
                old.rename(new)
                log(f"  [folders] renamed {old.name} → {new.name}")
            except OSError as e:
                log(f"[!] could not rename {old} → {new} ({e}); using the old folder")
                return old
        new.parent.mkdir(parents=True, exist_ok=True)
    return new


def src_dir(out_dir: Path, src: str) -> Path:
    """<out_dir>/<Arabic source title>/   (was <out_dir>/<slug>/)"""
    return _adopt(out_dir / src, out_dir / safe_name(src_title(src)))


def sub_dir(parent: Path, key: str) -> Path:
    """parent/<raw | md | docx | txt | json>  (renames an Arabic-named one from an older version)."""
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


# ═══════════════════════════════════════════════════════════════════════════
# HTTP
# ═══════════════════════════════════════════════════════════════════════════


def make_session(user_agent: str | None = None):
    s = http.Session()
    s.headers.update(HEADERS)
    if user_agent:
        s.headers["User-Agent"] = user_agent
    if _IMPERSONATE:
        s.impersonate = "firefox147" if "firefox" in s.headers["User-Agent"].lower() else "chrome124"
    return s


_net_lock = threading.Lock()
_net_requests = 0   # every HTTP request actually sent (retries included)


def net_requests() -> int:
    return _net_requests


def get_json(session, url: str, params: dict | None = None, referer: str | None = None,
             retries: int = 5, timeout: int = 25):
    """GET JSON with retry/backoff on 429 / 5xx / connection errors."""
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
                raise RuntimeError("HTTP 403 — blocked (try lowering --workers / raising --delay)")
            r.raise_for_status()
            return r.json()
        except RuntimeError:
            raise
        except Exception as e:  # network / JSON decode
            last = str(e)
            time.sleep(min(30, 2 ** attempt))
    raise RuntimeError(f"{url} failed after {retries} tries: {last}")


# ═══════════════════════════════════════════════════════════════════════════
# SOURCE METADATA  (which ayahs have material?)
# ═══════════════════════════════════════════════════════════════════════════


class Availability:
    """Wraps /sources/sources.json.  keys_flag ==  1 → data ONLY for ids in keys
                                    keys_flag == -1 → data for everything EXCEPT keys
                                    otherwise       → data everywhere."""

    def __init__(self, raw: dict | None):
        self.raw = raw or {}
        self._sets: dict[str, set[int]] = {}

    def has(self, src: str, ayah_abs_id: int) -> bool:
        meta = self.raw.get(src)
        if not isinstance(meta, dict):
            return True
        flag = meta.get("keys_flag")
        if flag not in (1, -1):
            return True
        ks = self._sets.get(src)
        if ks is None:
            ks = self._sets[src] = set(meta.get("keys") or [])
        return (ayah_abs_id in ks) if flag == 1 else (ayah_abs_id not in ks)


def merge_titles(data: dict) -> dict:
    """Write the Arabic name of every known source into sources.json ("title_ar" + group),
    adding an entry for any source the site's file doesn't list.  The site's own fields
    (keys / keys_flag / ...) are left untouched."""
    if not isinstance(data, dict):
        data = {}
    for slug, title in ALL_SOURCES.items():
        ent = data.get(slug)
        if not isinstance(ent, dict):
            ent = data[slug] = {}
        ent["title_ar"] = title
        grp = "tafsir" if slug in TAFSIR_SOURCES else ("extra" if slug in EXTRA_SOURCES else "uloom")
        ent.setdefault("group", grp)
    return data


def fetch_sources_meta(session, base: str, out_dir: Path) -> dict:
    cache = out_dir / "sources.json"
    try:
        data = get_json(session, f"{base}/sources/sources.json", {"json_rev": "1"})
    except Exception as e:
        log(f"[!] could not fetch sources.json ({e}); using cache / no pre-filter")
        data = load_json(cache, {}) or {}
    data = merge_titles(data)
    atomic_write_json(cache, data)
    return data


# ═══════════════════════════════════════════════════════════════════════════
# TEXT CLEANING  (mirrors the site's prep_data: footnotes [[…]], (p-N) markers)
# ═══════════════════════════════════════════════════════════════════════════

_TAG_RE = re.compile(r"<[^>]+>")
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


def md_color_verses(text: str) -> str:
    """Plain Markdown (--md_style plain) has no colour, so verses are **bold**."""
    out = []
    for seg, v in split_verses(text):
        if not v:
            out.append(seg)
        else:  # keep the ** markers hugging the text so Markdown accepts them
            lead = seg[:len(seg) - len(seg.lstrip())]
            trail = seg[len(seg.rstrip()):]
            out.append(f"{lead}**{seg.strip()}**{trail}")
    return "".join(out)


def md_poem(line: str) -> str:
    """Poetry line: the ؎ sign is dropped and the two hemistichs are separated by a gap
    (plain Markdown cannot centre or colour text; DOCX does)."""
    return poem_text(line)


def md_body(text: str) -> str:
    """One Markdown paragraph per source line (a single newline would be merged into
    one run-on paragraph by every Markdown renderer)."""
    return "\n\n".join(md_poem(l) if is_poem_line(l) else md_color_verses(l.strip())
                        for l in text.split("\n") if l.strip())


def md_note(text: str) -> str:
    return md_color_verses(" ".join(text.split()))


# ─── styled Markdown: inline HTML with the same colours / fonts / layout as the DOCX ───
# (set from --md_style; "plain" = pure Markdown without any HTML)
MD_STYLED = True
MD_FONT = "Kitab"          # --font
MD_SIZE = 16.0             # --docx_size (used as px)
MD_TITLE_SIZE = 20.0       # --tafsir_title_size
MD_TAFSIR_BOLD = False     # --tafsir_bold


def _hx(text: str) -> str:
    return _html.escape(text, quote=False)


def _fam(name: str) -> str:
    return "'" + name.replace("'", "") + "', serif"


def _st(color=None, font=None, size=None, bold=None, align=None, margin=None) -> str:
    p = []
    if color:
        p.append(f"color:#{color}")
    if font:
        p.append(f"font-family:{_fam(font)}")
    if size:
        p.append(f"font-size:{size:g}px")
    if bold is not None:
        p.append(f"font-weight:{'bold' if bold else 'normal'}")
    if align:
        p.append(f"text-align:{align}")
    if margin is not None:
        p.append(f"margin:{margin[0]}px 0 {margin[1]}px")
    return ";".join(p)


def _mdp(inner: str, color, font, size, bold=False, align=None, margin=None) -> str:
    return f'<p style="{_st(color, font, size, bold, align, margin)}">{inner}</p>'


def _mdh(level: int, text: str, color, font, size, bold: bool = True) -> str:
    return f'{"#" * level} <span style="{_st(color, font, size, bold)}">{_hx(text)}</span>'


def md_runs(text: str, base_color: str, sz: float) -> str:
    """Like _d_text_runs: Quran quotations in the verse font/colour, (١)-style numbers in
    the note colour, the rest inherits the paragraph's colour/font."""
    out = []
    for seg, is_verse in split_verses(text):
        if is_verse:
            out.append(f'<span style="{_st(VERSE_COLOR or base_color, MD_FONT, sz - VERSE_SIZE_DELTA, False)}">'
                       f'{_hx(seg)}</span>')
            continue
        for part in _NUM_MARK_RE.split(seg):
            if not part:
                continue
            if _NUM_MARK_RE.fullmatch(part):
                out.append(f'<span style="color:#{_note_color()}">{_hx(part)}</span>')
            else:
                out.append(_hx(part))
    return "".join(out)


def md_styled_body(body: str, tf: str, sz: float, ink: str, bold: bool = False) -> str:
    """Same line rules as the DOCX: poetry centred+red, basmala gold, '* * *' centred+red."""
    C = _DOCX_COLORS
    out = []
    for line in body.split("\n"):
        t = line.strip()
        if is_poem_line(line):
            out.append(_mdp(_hx(poem_text(line)), POEM_COLOR, tf, sz, align="center", margin=(4, 4)))
        elif t in BASMALA_LINES:
            bsz = sz + (2 if BASMALA_MODE == "text" else 10)
            out.append(_mdp(md_runs(t, "C99200", bsz), "C99200", tf, bsz, margin=(6, 6)))
        elif _SEP_LINE_RE.match(t):
            out.append(_mdp(md_runs(t, C["sep"], sz), C["sep"], tf, sz, align="center", margin=(6, 6)))
        elif t:
            out.append(_mdp(md_runs(t, ink, sz), ink, tf, sz, bold=bold, align="justify"))
    return "\n\n".join(out)


def md_styled_notes(notes: list[str], tf: str, sz: float) -> str:
    """Coloured rule, then «(١) text» with the number in the verse colour (as in the DOCX)."""
    if not notes:
        return ""
    nc = _note_color()
    # a plain <div> (not <hr>, which viewers decorate with their own extra line): short rule,
    # NOTE_RULE_CM wide, at the right-hand (start) side — like the DOCX
    rows = [f'<div style="width:{NOTE_RULE_CM:g}cm;height:0;border-top:1px solid #{nc};'
            f'margin:8px 0 6px;margin-left:auto;margin-right:0"></div>']
    for i, n in enumerate(notes, 1):
        rows.append(f'<p style="{_st(_DOCX_COLORS["note"], tf, sz, False, None, (2, 2))}">'
                    f'<span style="color:#{nc}">({ar_nums(i)}) </span>'
                    f'{md_runs(" ".join(n.split()), _DOCX_COLORS["note"], sz)}</p>')
    return "\n".join(rows)


def md_cover(title: str, title_font: str, title_extra: str, lines: list[str]) -> str:
    C = _DOCX_COLORS
    out = [_mdp(_hx(BASMALA), "C99200", MD_FONT, MD_SIZE + 2, align="center", margin=(0, 14)),
           f'<h1 style="{_st(C["title"], title_font, MD_SIZE + 14, True, "center")}">{_hx(title)}</h1>']
    if title_extra:
        out.append(_mdp(_hx(title_extra), C["surah"], MD_FONT, MD_SIZE + 2, align="center"))
    return "\n\n".join(out + lines)


def md_wrap(inner: str) -> str:
    """Whole file in one RTL block with the DOCX page colour, so the dark-ink text stays
    readable even in a dark viewer theme."""
    style = (f"background:#{PAGE_BG or 'FFFFFF'};color:#{_DOCX_COLORS['ink']};"
             f"{_st(font=TAFSIR_FONT, size=MD_SIZE)};line-height:{LINE_SPACING * 2:.2f};"
             f"padding:1.2em 1.6em;border-radius:8px")
    return f'<div dir="rtl" style="{style}">\n\n{inner}\n\n</div>\n'


def md_single(src: str, body: str, title_extra: str, plain_prefix: str = "") -> str:
    """Final text of a single-source .md file (cover + page wrapper when styled)."""
    if not MD_STYLED:
        return plain_prefix + body
    date = _mdp(_hx(f"المصدر: {BASE}/{src} | تاريخ التحميل: {time.strftime('%Y-%m-%d')}"),
                "888888", MD_FONT, 10, align="center", margin=(0, 12))
    return md_wrap(md_cover(src_title(src), TAFSIR_FONT, title_extra, [date]) + "\n\n" + body)


def md_styled_surah(src: str, surah: int, records: list[dict], quran) -> str:
    C = _DOCX_COLORS
    sz = MD_SIZE
    out = [_mdh(2, f"سورة {SUWAR[surah - 1][1]}", C["surah"], MD_FONT, sz + 6), ""]
    for r in sorted(records, key=lambda r: r["a_start"]):
        rng = (ar_nums(r["a_start"]) if r["a_start"] == r["a_end"]
               else f"{ar_nums(r['a_start'])}-{ar_nums(r['a_end'])}")
        out += [_mdh(3, f"الآية {rng}", C["ayah"], MD_FONT, sz + 2, bold=False), ""]
        head = quran_heading(quran, surah, r["a_start"], r["a_end"])
        if head:
            out += [_mdp(_hx(head), C["quran"], MD_FONT, sz, margin=(0, 6)), ""]
        body, notes = clean_text(r["data"])
        if not body:
            out += [_mdp("لا مواد لهذه الآية في هذا الكتاب", "888888", MD_FONT, sz - 2), ""]
            continue
        out += [md_styled_body(body, TAFSIR_FONT, sz, C["ink"]), ""]
        if notes:
            out += [md_styled_notes(notes, TAFSIR_FONT, max(10, sz - 4)), ""]
    return "\n".join(out)


def debug_marks(out_dir: Path, sources: list[str], surahs: list[int]) -> None:
    """Print unusual symbol characters found in saved data (to see what the site uses as quote marks)."""
    import unicodedata
    from collections import Counter
    cnt: Counter = Counter()
    sample: dict[str, str] = {}
    for src in sources:
        for s in surahs:
            for r in load_records(raw_paths(out_dir, src, s)[0]):
                d = r["data"]
                for i, ch in enumerate(d):
                    cat = unicodedata.category(ch)
                    if cat[0] in "PSC" and ch not in "\n .,:;!?-" and not (0x064B <= ord(ch) <= 0x065F):
                        cnt[ch] += 1
                        sample.setdefault(ch, d[max(0, i - 15): i + 25].replace("\n", " "))
    if not cnt:
        print("no data found (scrape something first)")
        return
    print("char  codepoint  count  name / example")
    for ch, n in cnt.most_common(40):
        print(f"{ch!r:6} U+{ord(ch):04X}  {n:6d}  {unicodedata.name(ch, '?')}  |  …{sample[ch]}…")


# ═══════════════════════════════════════════════════════════════════════════
# SCRAPING
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


def missing_ayahs(out_dir: Path, src: str, surah: int, a_lo: int, a_hi: int,
                  avail: Availability, prefilter: bool, force: bool) -> list[int]:
    """Ayahs that still need a network request, decided ONLY from what is already on disk
    (raw/*.jsonl, *.done) plus the cached sources.json.  Empty list = nothing to download."""
    jsonl, done = raw_paths(out_dir, src, surah)
    if force:
        return list(range(a_lo, a_hi + 1))
    if done.exists():
        return []
    covered: set[int] = set()
    for r in load_records(jsonl):
        covered.update(range(r["a_start"], r["a_end"] + 1))
    return [a for a in range(a_lo, a_hi + 1)
            if a not in covered and (not prefilter or avail.has(src, abs_id(surah, a)))]


def scrape_surah(session, base: str, out_dir: Path, src: str, surah: int,
                 a_lo: int, a_hi: int, avail: Availability, delay: float,
                 force: bool, prefilter: bool, partial: bool) -> dict:
    """Sequentially walk one surah for one source (jumping over grouped ayahs)."""
    jsonl, done = raw_paths(out_dir, src, surah)
    if force:
        if partial:   # only forget what overlaps the requested range; keep the rest of the surah
            recs0 = load_records(jsonl)
            hit = [r for r in recs0 if not (r["a_end"] < a_lo or r["a_start"] > a_hi)]
            if hit:   # a grouped block sticking out of the range is re-fetched whole
                a_lo, a_hi = min([a_lo] + [r["a_start"] for r in hit]), max([a_hi] + [r["a_end"] for r in hit])
            keep = [r for r in recs0 if r["a_end"] < a_lo or r["a_start"] > a_hi]
            jsonl.write_text("".join(json.dumps(r, ensure_ascii=False) + "\n" for r in keep),
                             encoding="utf-8")
        else:
            jsonl.write_text("", encoding="utf-8")
        done.unlink(missing_ok=True)
    elif done.exists():
        return {"src": src, "surah": surah, "status": "skipped", "fetched": 0}

    records = load_records(jsonl)
    covered: set[int] = set()
    for r in records:
        covered.update(range(r["a_start"], r["a_end"] + 1))

    n_ayahs = SUWAR[surah - 1][0]
    fetched = 0
    a = a_lo
    with open(jsonl, "a", encoding="utf-8") as fh:
        while a <= a_hi:
            if a in covered:
                a += 1
                continue
            if prefilter and not avail.has(src, abs_id(surah, a)):
                a += 1
                continue

            resp = get_json(
                session, f"{base}/get.php",
                {"src": src, "s": surah, "a": a, "ver": 1},
                referer=f"{base}/{src}/{surah}/{a}",
            )
            data = resp.get("data") if isinstance(resp, dict) else None
            data = data if isinstance(data, str) else ""

            start = int(resp.get("ayahs_start") or a) if isinstance(resp, dict) else a
            extra = int(resp.get("count") or 0) if isinstance(resp, dict) else 0
            end = min(n_ayahs, start + extra)
            start = min(start, a)  # never claim less than the requested ayah

            rec = {"src": src, "s": surah, "a_start": start, "a_end": max(end, a), "data": data}
            fh.write(json.dumps(rec, ensure_ascii=False) + "\n")
            fh.flush()
            records.append(rec)
            covered.update(range(rec["a_start"], rec["a_end"] + 1))
            fetched += 1
            a = rec["a_end"] + 1
            if delay:
                time.sleep(delay)

    if not partial:  # only mark complete when the whole surah was requested
        done.write_text(time.strftime("%Y-%m-%dT%H:%M:%S"), encoding="utf-8")
    return {"src": src, "surah": surah, "status": "done", "fetched": fetched}


# ═══════════════════════════════════════════════════════════════════════════
# OUTPUT
# ═══════════════════════════════════════════════════════════════════════════


def load_quran(session, base: str, out_dir: Path):
    path = out_dir / "quran.json"
    data = load_json(path)
    if data is None:
        data = get_json(session, f"{base}/sources/quran.json", {"json_rev": "1"})
        atomic_write_json(path, data)
    return data


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


def build_md_for_surah(src: str, surah: int, records: list[dict], quran) -> str:
    if MD_STYLED:
        return md_styled_surah(src, surah, records, quran)
    title = src_title(src)
    out = [f"## سورة {SUWAR[surah - 1][1]} — {title}", ""]
    for r in sorted(records, key=lambda r: r["a_start"]):
        head = quran_heading(quran, surah, r["a_start"], r["a_end"])
        rng = (ar_nums(r["a_start"]) if r["a_start"] == r["a_end"]
               else f"{ar_nums(r['a_start'])}-{ar_nums(r['a_end'])}")
        out.append(f"### الآية {rng}")
        out.append("")
        if head:
            out.append(f"**{head}**")
            out.append("")
        body, notes = clean_text(r["data"])
        out.append(md_body(body) if body else "_لا مواد لهذه الآية في هذا الكتاب_")
        out.append("")
        if notes:
            out.append("---")
            for i, n in enumerate(notes, 1):
                # two trailing spaces = hard line break: one footnote per line, no paragraph gap
                out.append(f"> ({ar_nums(i)}) {md_note(n)}" + ("  " if i < len(notes) else ""))
            out.append("")
    return "\n".join(out)


def _txt_for_surah(s: int, recs: list[dict]) -> str:
    lines = []
    for r in recs:
        body, notes = clean_text(r["data"])
        lines.append(f"[{SUWAR[s - 1][1]} {r['a_start']}-{r['a_end']}]\n{body}\n")
        lines.extend(f"({i}) {n}" for i, n in enumerate(notes, 1))
    return "\n".join(lines)


def _json_entries(recs: list[dict]) -> list[dict]:
    out = []
    for r in recs:
        body, notes = clean_text(r["data"])
        out.append({**r, "text": body, "footnotes": notes})
    return out


def build_outputs(out_dir: Path, src: str, surahs: list[int], args, quran) -> None:
    """Every requested format follows the same rule:
         (default)   one file per surah   -> <src>/<format>/<surah>.<ext>
         --merge     one file for all     -> <src>/<name>.<ext>"""
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

    if args.md:
        if not args.merge:
            for s, recs in sections:
                text = md_single(src, build_md_for_surah(src, s, recs, quran), f"سورة {SUWAR[s - 1][1]}")
                target("md", "md", s).write_text(text, encoding="utf-8")
        else:
            fm = f"# {title}\n\n> المصدر: {BASE}/{src}\n\n"
            body = "\n\n".join(build_md_for_surah(src, s, recs, quran) for s, recs in sections)
            text = md_single(src, body, "" if len(sections) == 1 else stem, plain_prefix=fm)
            target("md", "md").write_text(text, encoding="utf-8")
    if args.txt:
        if not args.merge:
            for s, recs in sections:
                target("txt", "txt", s).write_text(_txt_for_surah(s, recs), encoding="utf-8")
        else:
            target("txt", "txt").write_text("\n\n".join(_txt_for_surah(s, r) for s, r in sections),
                                          encoding="utf-8")
    if args.json:
        meta = {"source": src, "title": title, "origin": f"{BASE}/{src}"}
        if not args.merge:
            for s, recs in sections:
                atomic_write_json(target("json", "json", s), {**meta, "entries": _json_entries(recs)})
        else:
            atomic_write_json(target("json", "json"), {
                **meta, "entries": [e for _, recs in sections for e in _json_entries(recs)]})
    if args.docx:
        if not args.merge:
            for s, recs in sections:
                build_docx(target("docx", "docx", s), src, [(s, recs)], quran,
                           args.font, args.docx_size, title_extra=f"سورة {SUWAR[s - 1][1]}")
        else:
            build_docx(target("docx", "docx"), src, sections, quran, args.font, args.docx_size,
                       title_extra="" if len(sections) == 1 else stem)


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


def _d_sections(doc, src: str, sections, quran, font: str, size: float, page_break_surah: bool):
    """Emit surah → ayah-group content into *doc*.  sections = [(surah, records)]."""
    C = _DOCX_COLORS
    for idx, (s, recs) in enumerate(sections):
        _d_heading(doc, f"سورة {SUWAR[s - 1][1]}", 1, font, size + 6, C["surah"],
                   page_break=page_break_surah and idx > 0)
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
               title_extra: str = "") -> None:
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
    _d_par(doc, f"المصدر: {BASE}/{src} | تاريخ التحميل: {time.strftime('%Y-%m-%d')}",
           font, 10, color="888888", align=WD_ALIGN_PARAGRAPH.CENTER, after=12)

    _d_sections(doc, src, sections, quran, font, size, page_break_surah=True)
    doc.save(str(out_path))
    n = sum(len(r) for _, r in sections)
    log(f"  [docx] saved → {out_path}  ({len(sections)} surah(s), {n} entries)")


# ═══════════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════════


# ═══════════════════════════════════════════════════════════════════════════
# COMBINED OUTPUT  (--combine a,b,c : one file, per ayah → verse, then each tafsir in order)
# ═══════════════════════════════════════════════════════════════════════════


def resolve_combine(spec: str) -> list[str]:
    """'saadi,ibn-alqayim' -> ['saadi','ibn-alqayyim'] (order kept, typos auto-corrected)."""
    import difflib
    out: list[str] = []
    for raw in spec.split(","):
        slug = raw.strip()
        if not slug:
            continue
        if slug not in ALL_SOURCES:
            m = difflib.get_close_matches(slug, list(ALL_SOURCES), n=1, cutoff=0.8)
            if m:
                log(f"[!] --combine: '{slug}' not a known source, using '{m[0]}'")
                slug = m[0]
            else:
                log(f"[!] --combine: '{slug}' not in the built-in list (will still be tried)")
        if slug not in out:
            out.append(slug)
    return out


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


def md_styled_combined(order: list[str], surah: int, entries: list[dict], quran) -> str:
    C = _DOCX_COLORS
    sz = MD_SIZE
    out = [_mdh(2, f"سورة {SUWAR[surah - 1][1]}", C["surah"], MD_FONT, sz + 6), ""]
    for e in entries:
        out += [_mdh(3, f"الآية {_rng(e['a_start'], e['a_end'])}", C["ayah"], MD_FONT, sz + 2, bold=False), ""]
        head = quran_heading(quran, surah, e["a_start"], e["a_end"])
        if head:
            out += [_mdp(_hx(head), C["quran"], MD_FONT, sz, margin=(0, 6)), ""]
        for idx, src, rec in e["items"]:
            body, notes = clean_text(rec["data"])
            label = _src_label(idx, src)
            if (rec["a_start"], rec["a_end"]) != (e["a_start"], e["a_end"]):
                label += f" — الآيات {_rng(rec['a_start'], rec['a_end'])}"
            out += [_mdh(4, label, SRC_FG[(idx - 1) % len(SRC_FG)], TAFSIR_FONT, MD_TITLE_SIZE, bold=False), ""]
            out += [md_styled_body(body, TAFSIR_FONT, sz, C["ink"], bold=MD_TAFSIR_BOLD), ""]
            if notes:
                out += [md_styled_notes(notes, TAFSIR_FONT, max(10, sz - 4)), ""]
        for idx, src, rec in e["notes"]:
            out += [_mdp(_hx(_see_before(idx, src, rec)), "888888", MD_FONT, max(10, sz - 4),
                         margin=(0, 10)), ""]
    return "\n".join(out)


def md_combined_file(order: list[str], body: str, title_extra: str, plain_head: str) -> str:
    """Final text of a combined .md file (cover + page wrapper when styled)."""
    if not MD_STYLED:
        return plain_head + body
    names = [_mdp(_hx(_src_label(i, src)), SRC_FG[(i - 1) % len(SRC_FG)], TAFSIR_FONT, MD_TITLE_SIZE,
                  align="center", margin=(0, 0)) for i, src in enumerate(order, 1)]
    date = _mdp(_hx(f"تاريخ التحميل: {time.strftime('%Y-%m-%d')}"), "888888", MD_FONT, 10,
                align="center", margin=(8, 12))
    return md_wrap(md_cover("جمع التفاسير", MD_FONT, title_extra, names + [date]) + "\n\n" + body)


def build_combined_md(order: list[str], surah: int, entries: list[dict], quran) -> str:
    if MD_STYLED:
        return md_styled_combined(order, surah, entries, quran)
    out = [f"## سورة {SUWAR[surah - 1][1]}", ""]
    for e in entries:
        out += [f"### الآية {_rng(e['a_start'], e['a_end'])}", ""]
        head = quran_heading(quran, surah, e["a_start"], e["a_end"])
        if head:
            out += [f"**{head}**", ""]
        for idx, src, rec in e["items"]:
            body, notes = clean_text(rec["data"])
            label = _src_label(idx, src)
            if (rec["a_start"], rec["a_end"]) != (e["a_start"], e["a_end"]):
                label += f" — الآيات {_rng(rec['a_start'], rec['a_end'])}"
            out += [f"#### {label}", ""]
            out += [md_body(body), ""]
            for i, n in enumerate(notes, 1):
                # two trailing spaces = hard line break: one footnote per line, no paragraph gap
                out.append(f"> ({ar_nums(i)}) {md_note(n)}" + ("  " if i < len(notes) else ""))
            if notes:
                out.append("")
        for idx, src, rec in e["notes"]:
            out += [f"_{_see_before(idx, src, rec)}_", ""]
    return "\n".join(out)


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
                        bold: bool = False) -> None:
    """sections = [(surah, entries)] from combined_entries()."""
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
    _d_par(doc, f"تاريخ التحميل: {time.strftime('%Y-%m-%d')}", font, 10, color="888888",
           align=WD_ALIGN_PARAGRAPH.CENTER, before=8, after=12)

    for k, (s, entries) in enumerate(sections):
        _d_heading(doc, f"سورة {SUWAR[s - 1][1]}", 1, font, size + 6, C["surah"],
                   page_break=k > 0)
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
    if args.md:
        names = "، ".join(_src_label(i, src) for i, src in enumerate(order, 1))
        head = f"# جمع التفاسير\n\n> {names}\n\n"
        if not args.merge:
            md_dir = sub_dir(root, "md")
            for s, entries in sections:
                md = build_combined_md(order, s, entries, quran)
                text = md_combined_file(order, md, f"سورة {SUWAR[s - 1][1]}", head)
                (md_dir / f"{surah_file(s)}.md").write_text(text, encoding="utf-8")
            log(f"  [md] saved → {md_dir}  ({len(sections)} surah file(s))")
        else:
            mds = [build_combined_md(order, s, entries, quran) for s, entries in sections]
            text = md_combined_file(order, "\n\n".join(mds), "" if len(sections) == 1 else stem, head)
            (root / f"{stem}.md").write_text(text, encoding="utf-8")
            log(f"  [md] saved → {root / (stem + '.md')}")
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
                                title_size=args.tafsir_title_size, bold=args.tafsir_bold)


def report_incomplete(out_dir: Path, sources: list[str], surahs: list[int]) -> None:
    """Warn (before building) about surahs that were never fully downloaded, so a merged
    file is never silently incomplete."""
    for src in sources:
        missing = [s for s in surahs if not raw_paths(out_dir, src, s)[1].exists()]
        if missing:
            shown = ", ".join(str(s) for s in missing[:20]) + (" …" if len(missing) > 20 else "")
            log(f"  [!] {src_title(src)}: {len(missing)} surah(s) not marked complete "
                f"(no .done): {shown} — the output may be partial; re-run the same command "
                f"to download what is missing")


def print_status(out_dir: Path, sources: list[str], surahs: list[int],
                 avail: Availability, prefilter: bool) -> None:
    """Progress from disk only (no network)."""
    for src in sources:
        done = partial = todo = 0
        for s in surahs:
            jsonl, d = raw_paths(out_dir, src, s)
            if d.exists():
                done += 1
            elif jsonl.exists() and jsonl.stat().st_size > 0:
                partial += 1
            if missing_ayahs(out_dir, src, s, 1, SUWAR[s - 1][0], avail, prefilter, False):
                todo += 1
        print(f"  {src:24s} done {done:3d}/{len(surahs)}   partial {partial:3d}   "
              f"need download {todo:3d}")


def download_all(session, args, out_dir: Path, todo: list, avail: Availability,
                 partial: bool) -> tuple[int, int, int]:
    """Run scrape_surah for every pending (src, surah) on a thread pool → (ok, skipped, failed)."""
    pool = ThreadPoolExecutor(max_workers=args.workers)
    futs = {
        pool.submit(scrape_surah, session, args.base, out_dir, src, s, lo, hi, avail,
                    args.delay, args.force, not args.no_prefilter, partial): (src, s)
        for src, s, lo, hi, _n in todo
    }
    ok = skipped = failed = 0
    try:
        for i, f in enumerate(as_completed(futs), 1):
            src, s = futs[f]
            try:
                res = f.result()
                if res["status"] == "skipped":
                    skipped += 1
                else:
                    ok += 1
                log(f"  [{i}/{len(todo)}] {src} / {s:03d} {SUWAR[s-1][1]}: "
                    f"{res['status']} ({res['fetched']} requests)")
            except Exception as e:
                failed += 1
                log(f"  [{i}/{len(todo)}] {src} / {s:03d}: ERROR {e}")
    except KeyboardInterrupt:
        log("\n[!] Interrupted — cancelling queued work. Progress saved; re-run to resume.")
        pool.shutdown(wait=False, cancel_futures=True)
        sys.exit(1)
    pool.shutdown(wait=True)
    return ok, skipped, failed


def main():
    p = argparse.ArgumentParser(description="Scrape tafsir.app sources to JSONL / JSON / Markdown")
    p.add_argument("--src", default=None, help="comma-separated source slugs, e.g. saadi,muyassar")
    p.add_argument("--group", choices=list(GROUPS), default=None,
                   help="add every source of a tab: tafsir | uloom")
    p.add_argument("--surah", default=None, help="'all', '53', or '1-3,53,112-114'")
    p.add_argument("--ayahs", default=None, help="ayah range inside ONE surah, e.g. 33-54")
    p.add_argument("--out_dir", default="./tafsir_output")
    p.add_argument("--workers", type=int, default=4, help="parallel (source, surah) tasks")
    p.add_argument("--delay", type=float, default=0.3, help="pause after each request (per worker)")
    p.add_argument("--force", action="store_true",
                   help="re-download even if already saved (with --ayahs: only that range). "
                        "Without it, saved data is never requested again")
    p.add_argument("--no_prefilter", action="store_true",
                   help="request every ayah even if sources.json says there's no data")
    p.add_argument("--base", default=BASE, help=argparse.SUPPRESS)
    p.add_argument("--user-agent", "--ua", dest="ua", default=None)

    p.add_argument("--md", action=argparse.BooleanOptionalAction, default=None,
                   help="Markdown output. With no format flag at all, md + docx are built; as soon "
                        "as you pass any of --md/--docx/--json/--txt, ONLY the formats you named "
                        "are built. --no-md removes it from the default pair (also for --combine)")
    p.add_argument("--json", action="store_true", help="JSON output (entries with cleaned text and footnotes)")
    p.add_argument("--txt", action="store_true", help="plain-text output")
    p.add_argument("--docx", action=argparse.BooleanOptionalAction, default=None,
                   help="Word output (needs python-docx). Same rule as --md: built by default "
                        "only when no format flag is given; --no-docx removes it from the default pair")
    p.add_argument("--merge", action="store_true",
                   help="put all requested surahs in ONE file per format (default: one file per "
                        "surah for every format; also applies to --combine)")
    p.add_argument("--font", default="Kitab",
                   help="DOCX font (must be installed on the machine opening the file). Word has no CSS-style "
                        "fallback stack, so only the first name of a list like \"Kitab, system-ui\" is used")
    p.add_argument("--docx_size", type=float, default=16, help="DOCX body size in pt (default 16)")
    p.add_argument("--tafsir_font", default="KFGQPC Uthman Taha Naskh",
                   help="DOCX font for tafsir names and tafsir text (default 'KFGQPC Uthman Taha Naskh'); "
                        "Quran verses (﴿…﴾ heading and quotations) stay in --font")
    p.add_argument("--tafsir_title_size", type=float, default=None,
                   help="--combine only: DOCX size in pt of the tafsir names (default 20)")
    p.add_argument("--tafsir_bold", action=argparse.BooleanOptionalAction, default=None,
                   help="--combine only: bold tafsir text in DOCX files (default off; "
                        "--tafsir_bold to enable)")
    p.add_argument("--line_spacing", type=float, default=0.9,
                   help="DOCX line spacing multiple (default 0.9; 1.0 = font's natural height, 1.25 = airy)")
    p.add_argument("--page_bg", default="F3EEE2",
                   help="hex colour for the WHOLE DOCX page background (default F3EEE2, soft "
                        "paper); use 'none' for a white page")
    p.add_argument("--verse_color", default="0F6B9C",
                   help="hex colour for ﴿…﴾ quotations inside the tafsir text (DOCX and styled Markdown). "
                        "Default 0F6B9C; use 'none' to disable")
    p.add_argument("--with_quran", action=argparse.BooleanOptionalAction, default=True,
                   help="print the ﴿ayah﴾ heading above each entry (default on; "
                        "--no-with_quran to disable)")
    p.add_argument("--md_style", choices=["styled", "plain"], default="styled",
                   help="styled (default): Markdown files carry the same colours, fonts, centred "
                        "poetry, footnote rule and page background as the Word files, using inline "
                        "HTML (renders in VS Code, Typora, browsers; GitHub strips the styles). "
                        "plain: pure Markdown, verses in **bold**, no HTML")
    p.add_argument("--ligatures", choices=["auto", "expand", "keep"], default="auto",
                   help="honorific ligatures like ﵇ (U+FD47) that many fonts can't draw: "
                        "auto = spell them out but keep ﷺ (default), expand = spell out all "
                        "incl. ﷺ, keep = leave as-is")
    p.add_argument("--verse_heuristic", action=argparse.BooleanOptionalAction, default=True,
                   help="also colour vowelled quotes inside ( ) « » [ ] (default on)")
    p.add_argument("--debug_marks", action="store_true",
                   help="list unusual symbols in the saved data (to see which quote marks the site uses) and exit")
    p.add_argument("--combine", default=None, metavar="SLUGS",
                   help="comma-separated sources, IN ORDER, merged into ONE file: per ayah the "
                        "verse, then each tafsir in that order → "
                        "<out_dir>/جمع_التفاسير/<مصدر١>_<مصدر٢>/ (one folder per combination). Sources are scraped if "
                        "missing, so it works while downloading or after.")
    p.add_argument("--combine_only", action="store_true",
                   help="requires --combine: build ONLY the combined files (skip the "
                        "individual per-source outputs; --json/--txt are then ignored)")
    p.add_argument("--sources_file", default=None,
                   help="extra JSON of Arabic source titles (default: sources_ar.json next to the script)")
    p.add_argument("--list_sources", action="store_true", help="list known source slugs and exit")
    p.add_argument("--status", action="store_true", help="show progress and exit")
    args = p.parse_args()
    args.font = args.font.split(",")[0].strip() or "Kitab"   # Word takes ONE font name, not a CSS stack

    # ── output formats: no flag → md + docx; any positive format flag → ONLY those ──
    if args.md or args.docx or args.json or args.txt:
        args.md, args.docx = bool(args.md), bool(args.docx)
    else:
        args.md = True if args.md is None else args.md
        args.docx = True if args.docx is None else args.docx

    # ── flag coherence ────────────────────────────────────────────────────
    if args.combine_only and not args.combine:
        p.error("--combine_only needs --combine SLUGS")
    if args.combine is None and (args.tafsir_title_size is not None or args.tafsir_bold is not None):
        log("[!] --tafsir_title_size / --tafsir_bold only affect --combine DOCX files — ignored")
    args.tafsir_title_size = 20 if args.tafsir_title_size is None else args.tafsir_title_size
    args.tafsir_bold = bool(args.tafsir_bold)
    if args.combine and (args.json or args.txt):
        log("[!] --json/--txt are only built for the individual per-source files, "
            "not for --combine (combined output is md/docx only)"
            + (" — with --combine_only they are ignored" if args.combine_only else ""))
    if args.docx and _docxlib is None:
        log("[!] python-docx not installed — DOCX output disabled (pip install python-docx)")
        args.docx = False
    if not (args.md or args.docx or args.json or args.txt):
        log("[!] no output format enabled: the script will only download "
            "(add --md, --docx, --json or --txt to build files)")
    if args.combine and not (args.md or args.docx):
        log("[!] --combine files are md/docx only: add --md and/or --docx to build them")

    global VERSE_COLOR, VERSE_HEURISTIC, LIGATURES, PAGE_BG, LINE_SPACING, TAFSIR_FONT
    LINE_SPACING = args.line_spacing
    TAFSIR_FONT = args.tafsir_font.split(",")[0].strip() or "KFGQPC Uthman Taha Naskh"
    PAGE_BG = None if args.page_bg.lower() in ("none", "off", "white", "") else args.page_bg.lstrip("#").upper()
    LIGATURES = args.ligatures
    global MD_STYLED, MD_FONT, MD_SIZE, MD_TITLE_SIZE, MD_TAFSIR_BOLD
    MD_STYLED = args.md_style == "styled"
    MD_FONT, MD_SIZE = args.font, float(args.docx_size)
    MD_TITLE_SIZE, MD_TAFSIR_BOLD = float(args.tafsir_title_size), bool(args.tafsir_bold)
    VERSE_HEURISTIC = args.verse_heuristic
    VERSE_COLOR = None if args.verse_color.lower() in ("none", "off", "") else args.verse_color.lstrip("#").upper()

    load_source_registry(args.sources_file)

    if args.list_sources:
        for grp, d in GROUPS.items():
            print(f"[{grp}]")
            for k, v in d.items():
                print(f"  {k:24s} {v}")
        print("[extra — non-text, pass explicitly with --src]")
        for k, v in EXTRA_SOURCES.items():
            print(f"  {k:24s} {v}")
        return

    sources: list[str] = []
    if args.group:
        sources += list(GROUPS[args.group])
    if args.src:
        sources += [s.strip() for s in args.src.split(",") if s.strip()]
    combine_order = resolve_combine(args.combine) if args.combine else []
    sources += combine_order
    sources = list(dict.fromkeys(sources))
    if not sources or not args.surah:
        p.error("need --src and/or --group, plus --surah (see --list_sources)")
    unknown = [s for s in sources if s not in ALL_SOURCES]
    if unknown:
        log(f"[!] not in the built-in list (will still be tried): {', '.join(unknown)}")

    surahs = parse_surah_spec(args.surah)
    if args.ayahs and len(surahs) != 1:
        p.error("--ayahs only works with a single --surah")
    if args.ayahs:
        parse_ayah_range(args.ayahs, surahs[0])   # validate early

    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    prefilter = not args.no_prefilter
    # cached sources.json (from an earlier run) lets us decide what is missing WITHOUT the network
    avail = Availability((load_json(out_dir / "sources.json", {}) or None) if prefilter else None)

    if args.status:
        print_status(out_dir, sources, surahs, avail, prefilter)
        return
    if args.debug_marks:
        debug_marks(out_dir, sources, surahs)
        return

    partial = bool(args.ayahs)
    session = None

    def get_session():
        nonlocal session
        if session is None:
            session = make_session(args.ua)
        return session

    # ── 1. plan, from disk only: which (source, surah) really need the network? ──
    todo = []       # (src, surah, lo, hi, n_missing_ayahs)
    complete = 0
    for src in sources:
        for s in surahs:
            lo, hi = parse_ayah_range(args.ayahs, s)
            miss = missing_ayahs(out_dir, src, s, lo, hi, avail, prefilter, args.force)
            if miss:
                todo.append((src, s, lo, hi, len(miss)))
                continue
            complete += 1
            if not partial:   # everything is saved but the marker is missing → repair it
                jsonl, done = raw_paths(out_dir, src, s)
                if not done.exists():
                    jsonl.touch()
                    done.write_text(time.strftime("%Y-%m-%dT%H:%M:%S"), encoding="utf-8")

    total = len(sources) * len(surahs)
    if todo:
        log(f"[*] {total} (source, surah) pair(s): {complete} already on disk, {len(todo)} to "
            f"download (up to {sum(t[4] for t in todo)} ayah(s) not saved yet), {args.workers} workers")
    else:
        log(f"[*] {total} (source, surah) pair(s): everything is already on disk — nothing to download")

    # ── 2. network, only if something is missing ──
    ok = failed = 0
    skipped = complete
    if todo:
        if args.no_prefilter:
            atomic_write_json(out_dir / "sources.json",
                              merge_titles(load_json(out_dir / "sources.json", {}) or {}))
        else:
            avail = Availability(fetch_sources_meta(get_session(), args.base, out_dir))

    quran = None
    if args.with_quran:
        quran = load_json(out_dir / "quran.json")
        if not quran:   # the only request a fully-downloaded run can still make (once, then cached)
            try:
                quran = load_quran(get_session(), args.base, out_dir)
            except Exception as e:
                log(f"[!] could not load quran.json ({e}) — continuing without ﴿ayah﴾ headings")

    if todo:
        ok, sk, failed = download_all(get_session(), args, out_dir, todo, avail, partial)
        skipped += sk

    if not partial:
        report_incomplete(out_dir, sources, surahs)

    log("[*] Building outputs ...")
    if not (combine_order and args.combine_only):
        for src in sources:
            build_outputs(out_dir, src, surahs, args, quran)
    if combine_order:
        build_combined(out_dir, combine_order, surahs, args, quran)

    n_req = net_requests()
    log(f"[✓] finished: downloaded={ok} already-on-disk={skipped} failed={failed} | "
        f"network requests: {n_req}" + (" (everything was already on disk)" if n_req == 0 else "")
        + f" → {out_dir}")
    if failed:
        log("    re-run the same command to retry failed (source, surah) tasks.")


if __name__ == "__main__":
    main()