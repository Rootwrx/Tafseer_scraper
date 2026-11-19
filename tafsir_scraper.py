#!/usr/bin/env python3
"""
tafsir_scraper.py — Scrape tafsir / ulum sources from tafsir.app (الباحث القرآني)
and build Word (.docx) files from them (raw data is kept as JSONL).

Files  (keep the three together in one folder; run only this one)
─────
    tafsir_scraper.py   this file: command line + downloading (network, resume, progress)
    tafsir_docx.py      builds the Word files from the saved raw data (no network)
    tafsir_common.py    shared: surah table, source names, folder layout, raw-data helpers

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

Output tree   (same layout as transform.py:  <group>/<category>/<source>)
───────────
    <out_dir>/تفسير/أمّهات/تفسير ابن كثير/raw/surah_053.jsonl   raw records (source of truth)
    <out_dir>/تفسير/أمّهات/تفسير ابن كثير/raw/surah_053.done    completion marker
    <out_dir>/علوم/غريب ومعاني/معاني الزجاج/...                 (the "uloom" tab)
    (Word files are built ONLY when asked for — no --docx = download only)
    <SRC>/docx/سورة النجم.docx    with --docx  per-surah Word
      where <SRC> = <out_dir>/<group>/<category>/<عنوان المصدر>
    Sources outside tafsir/uloom (e.g. iraab-graphs) stay at <out_dir>/<عنوان المصدر>/.
    --layout flat  → the old  <out_dir>/<عنوان المصدر>/  layout (no group/category folders).
    --merge → ONE file directly inside <SRC>/ :
        سورة الفاتحة إلى سورة الناس.docx | القرآن كاملا.docx | سورة النجم إلى سورة الرحمن.docx
        سور 1-3،53،112-114.docx          (non-contiguous selections are named by number)
    <out_dir>/جمع_التفاسير/<مصدر١>_<مصدر٢>/docx/سورة النجم.docx   with --combine
    (old flat <slug>/ or <title>/ folders are moved into <group>/<category>/<title> automatically)
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

    # NOTE: without --docx the script ONLY downloads (raw/ data).

    # Word output (RTL, real Heading styles → Navigation Pane)
    python tafsir_scraper.py --src saadi --surah 53 --docx
    python tafsir_scraper.py --src saadi --surah 1-3 --docx --font "Kitab" --docx_size 18

    # ONE merged file instead of one file per surah
    python tafsir_scraper.py --src ibn-katheer --surah all --docx --merge
    python tafsir_scraper.py --src ibn-katheer --surah 1-9 --docx --merge
    #   (merged --docx files also get the book intro from get_info.php, cached in <out_dir>/info/;
    #    --no_info turns that off; per-surah files never get it)
    #   already downloaded?  run the SAME command again: it rebuilds from the saved raw/
    #   data and only requests what is missing (nothing missing → zero network requests)

    # combine several tafsirs into ONE docx (order = order given)
    #   per ayah: ﴿verse﴾ → saadi → ibn-alqayyim → …
    python tafsir_scraper.py --combine saadi,ibn-alqayyim,muyassar --surah 53 --docx
    python tafsir_scraper.py --combine saadi,muyassar --surah 1-3 --docx --combine_only
    python tafsir_scraper.py --combine saadi,muyassar --surah all --docx --merge --combine_only
    #   → <out_dir>/جمع_التفاسير/<مصدر١>_<مصدر٢>/docx/سورة النجم.docx

    # extras
    python tafsir_scraper.py --list_sources
    #   an unknown source name is rejected up front ("did you mean …"); --allow_unknown overrides
    python tafsir_scraper.py --src saadi --surah 53 --force      # re-download even if saved
    python tafsir_scraper.py --src saadi --surah all --status    # progress only (no network)

Install:  pip install curl_cffi python-docx   (curl_cffi optional; python-docx only for --docx)
"""

import argparse
import json
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

import tafsir_common as _common          # LAYOUT is set on this module in main()
from tafsir_common import (
    ALL_SOURCES, BASE, EXTRA_SOURCES, GROUPS, SUWAR, TAFSIR_SOURCES, _clean_intro, abs_id,
    atomic_write_json, load_json, load_records, load_source_registry, log, parse_ayah_range,
    parse_categories_html, parse_surah_spec, raw_paths, set_categories, set_site_meta, src_title,
)
from tafsir_docx import (
    SOURCE_INTRO, _docxlib, build_combined, build_outputs, configure as configure_docx,
)

HEADERS = {
    "User-Agent": "Mozilla/5.0 (X11; Linux x86_64; rv:128.0) Gecko/20100101 Firefox/128.0",
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "ar,en-US;q=0.5",
}

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


def validate_sources(get_session, base: str, out_dir: Path, sources: list[str],
                     allow_unknown: bool) -> None:
    """Stop BEFORE any request / folder / file is made for a slug that does not exist.

    An unknown slug used to be "still tried": get.php answers with empty data for every
    ayah, and the script saved those empty records and marked the surah .done as if it
    were a real source.  A slug is accepted when it is in the built-in lists /
    sources_ar.json, or listed in the site's own sources.json (cached copy first, then a
    fresh download — at most ONE request, and only when something is not known locally).
    """
    import difflib
    cache = out_dir / "sources.json"
    cached = load_json(cache, {}) or {}
    unknown = [s for s in sources if s not in ALL_SOURCES and s not in cached]
    if not unknown:
        return
    site = None
    try:
        site = get_json(get_session(), f"{base}/sources/sources.json", {"json_rev": "1"})
    except Exception as e:
        log(f"[!] could not download sources.json to verify: {', '.join(unknown)} ({e})")
    listed = set(site) if isinstance(site, dict) and site else None
    if listed is None:
        if allow_unknown:
            log("[!] cannot verify the source name(s) — continuing because of --allow_unknown")
            return
        raise SystemExit(
            f"[!] cannot verify source(s): {', '.join(unknown)} — not in the built-in list and "
            f"sources.json could not be downloaded.\n    Check the spelling and your connection, "
            f"or pass --allow_unknown to try anyway.")
    atomic_write_json(cache, merge_titles(site))      # refresh the cache while we have it
    bad = [s for s in unknown if s not in listed]
    if not bad:
        return
    cands = sorted(set(ALL_SOURCES) | listed)
    lines = []
    for s in bad:
        near = difflib.get_close_matches(s, cands, n=3, cutoff=0.6)
        names = ", ".join(f"{x} ({src_title(x)})" for x in near)
        lines.append(f"    {s}" + (f"   — did you mean: {names}?" if near else ""))
    if allow_unknown:
        log("[!] source(s) not found on the site — continuing because of --allow_unknown:\n"
            + "\n".join(lines))
        return
    raise SystemExit(
        "[!] unknown source(s) — they do not exist on tafsir.app, nothing was downloaded:\n"
        + "\n".join(lines)
        + "\n    (--list_sources shows the known slugs; --allow_unknown skips this check)")


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


def has_material(avail: Availability, src: str, surah: int, prefilter: bool,
                 lo: int = 1, hi: int | None = None) -> bool:
    """False when sources.json says this source has NO material for any ayah in lo..hi
    (e.g. a source that only covers some surahs).  Such surahs are skipped completely:
    no request, no empty .jsonl, no .done marker."""
    if not prefilter:
        return True
    hi = SUWAR[surah - 1][0] if hi is None else hi
    return any(avail.has(src, abs_id(surah, a)) for a in range(lo, hi + 1))


def scrape_surah(session, base: str, out_dir: Path, src: str, surah: int,
                 a_lo: int, a_hi: int, avail: Availability, delay: float,
                 force: bool, prefilter: bool, partial: bool) -> dict:
    """Sequentially walk one surah for one source (jumping over grouped ayahs)."""
    if not has_material(avail, src, surah, prefilter, a_lo, a_hi):
        return {"src": src, "surah": surah, "status": "no material", "fetched": 0}
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
            if not isinstance(resp, dict) or "data" not in resp:
                # an error / unexpected JSON must never be stored as "this ayah is empty"
                raise RuntimeError(f"unexpected reply for {src} {surah}:{a} "
                                   f"({str(resp)[:100]!r}) — not saved")
            data = resp["data"] if isinstance(resp["data"], str) else ""

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
        if prefilter and records and all(not (r.get("data") or "").strip() for r in records):
            # sources.json promised material here, yet every reply was empty → something is
            # wrong (bad / removed source, blocked).  Keep NOTHING so a re-run asks again.
            jsonl.write_text("", encoding="utf-8")
            return {"src": src, "surah": surah, "status": "empty", "fetched": fetched}
        done.write_text(time.strftime("%Y-%m-%dT%H:%M:%S"), encoding="utf-8")
    return {"src": src, "surah": surah, "status": "done", "fetched": fetched}


# ═══════════════════════════════════════════════════════════════════════════
# QURAN TEXT  (for the ﴿ayah﴾ headings)
# ═══════════════════════════════════════════════════════════════════════════


def load_quran(session, base: str, out_dir: Path):
    path = out_dir / "quran.json"
    data = load_json(path)
    if data is None:
        data = get_json(session, f"{base}/sources/quran.json", {"json_rev": "1"})
        atomic_write_json(path, data)
    return data


def fetch_source_info(session, base: str, out_dir: Path, src: str, force: bool = False) -> str:
    """Intro text of a book from  <base>/get_info.php?src=<src>  ('' if the book has none).
    Cached in <out_dir>/info/<src>.json, so it is requested at most once per book.
    A book without info (empty / 4xx answer) is cached as empty; a network failure is not."""
    global _net_requests
    path = out_dir / "info" / f"{src}.json"
    if not force:
        cached = load_json(path)
        if isinstance(cached, dict):
            return _clean_intro(cached.get("data"))
    data, ok, last = {}, False, ""
    for attempt in range(1, 4):
        try:
            with _net_lock:
                _net_requests += 1
            r = session.get(f"{base}/get_info.php", params={"src": src},
                            headers={"Referer": f"{base}/{src}"}, timeout=25)
            if r.status_code == 429 or r.status_code >= 500:
                last = f"HTTP {r.status_code}"
                time.sleep(2 ** attempt)
                continue
            if r.status_code == 403:
                last = "HTTP 403"
                break
            if r.status_code < 400:
                try:
                    data = r.json()
                except ValueError:
                    data = {}
            ok = True          # 200 or a 4xx like 404 → the book simply has no info
            break
        except Exception as e:
            last = str(e)
            time.sleep(2 ** attempt)
    if not ok:
        log(f"  [!] intro of {src} not fetched ({last}) — continuing without it")
        return ""
    raw = data.get("data") if isinstance(data, dict) else ""
    path.parent.mkdir(parents=True, exist_ok=True)
    atomic_write_json(path, {"data": raw if isinstance(raw, str) else ""})
    return _clean_intro(raw)


# ═══════════════════════════════════════════════════════════════════════════
# --combine a,b,c  (slug resolution; the combined Word files are built in tafsir_docx.py)
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


def report_incomplete(out_dir: Path, sources: list[str], surahs: list[int],
                      avail: Availability, prefilter: bool) -> None:
    """Warn (before building) about surahs that were never fully downloaded, so a merged
    file is never silently incomplete."""
    for src in sources:
        missing = [s for s in surahs if has_material(avail, src, s, prefilter)
                   and not raw_paths(out_dir, src, s)[1].exists()]
        if missing:
            shown = ", ".join(str(s) for s in missing[:20]) + (" …" if len(missing) > 20 else "")
            log(f"  [!] {src_title(src)}: {len(missing)} surah(s) not marked complete "
                f"(no .done): {shown} — the output may be partial; re-run the same command "
                f"to download what is missing")


def print_status(out_dir: Path, sources: list[str], surahs: list[int],
                 avail: Availability, prefilter: bool) -> None:
    """Progress from disk only (no network)."""
    for src in sources:
        done = partial = todo = nomat = 0
        for s in surahs:
            if not has_material(avail, src, s, prefilter):
                nomat += 1
                continue
            jsonl, d = raw_paths(out_dir, src, s)
            if d.exists():
                done += 1
            elif jsonl.exists() and jsonl.stat().st_size > 0:
                partial += 1
            if missing_ayahs(out_dir, src, s, 1, SUWAR[s - 1][0], avail, prefilter, False):
                todo += 1
        print(f"  {src:24s} done {done:3d}/{len(surahs)}   partial {partial:3d}   "
              f"need download {todo:3d}   no material {nomat:3d}")


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
                if res["status"] == "empty":
                    failed += 1
                    log(f"  [{i}/{len(todo)}] {src} / {s:03d} {SUWAR[s-1][1]}: EMPTY — every reply "
                        f"had no text, nothing saved (wrong or removed source?)")
                    continue
                if res["status"] in ("skipped", "no material"):
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
    p = argparse.ArgumentParser(description="Scrape tafsir.app sources to JSONL and build Word (.docx) files")
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
    p.add_argument("--allow_unknown", action="store_true",
                   help="do not stop on a source name that is neither built-in nor listed in the "
                        "site's sources.json (default: unknown names are rejected up front)")
    p.add_argument("--layout", choices=["grouped", "flat"], default="grouped",
                   help="grouped (default): <out_dir>/<group>/<category>/<source>/ (same as "
                        "transform.py); flat: the old <out_dir>/<source>/")
    p.add_argument("--html", default=None, metavar="FILE",
                   help="saved tafsir.app page: read the category tabs from it instead of the "
                        "built-in table")
    p.add_argument("--base", default=BASE, help=argparse.SUPPRESS)
    p.add_argument("--user-agent", "--ua", dest="ua", default=None)

    p.add_argument("--docx", action="store_true",
                   help="build Word files (needs python-docx). Output is explicit: without "
                        "--docx the script only downloads the raw data")
    p.add_argument("--merge", action="store_true",
                   help="put all requested surahs in ONE file per format (default: one file per "
                        "surah for every format; also applies to --combine)")
    p.add_argument("--no_info", action="store_true",
                   help="do not add the book intro (get_info.php) to merged docx files")
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
                   help="hex colour for ﴿…﴾ quotations inside the tafsir text (DOCX). "
                        "Default 0F6B9C; use 'none' to disable")
    p.add_argument("--with_quran", action=argparse.BooleanOptionalAction, default=True,
                   help="print the ﴿ayah﴾ heading above each entry in the docx (default on; "
                        "--no-with_quran to disable). quran.json is only fetched when "
                        "--docx is requested")
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
                        "individual per-source outputs)")
    p.add_argument("--sources_file", default=None,
                   help="extra JSON of Arabic source titles (default: sources_ar.json next to the script)")
    p.add_argument("--list_sources", action="store_true", help="list known source slugs and exit")
    p.add_argument("--status", action="store_true", help="show progress and exit")
    args = p.parse_args()
    args.font = args.font.split(",")[0].strip() or "Kitab"   # Word takes ONE font name, not a CSS stack

    # ── flag coherence ────────────────────────────────────────────────────
    if args.combine_only and not args.combine:
        p.error("--combine_only needs --combine SLUGS")
    if args.combine is None and (args.tafsir_title_size is not None or args.tafsir_bold is not None):
        log("[!] --tafsir_title_size / --tafsir_bold only affect --combine DOCX files — ignored")
    args.tafsir_title_size = 20 if args.tafsir_title_size is None else args.tafsir_title_size
    args.tafsir_bold = bool(args.tafsir_bold)
    if args.docx and _docxlib is None:
        log("[!] python-docx not installed — DOCX output disabled (pip install python-docx)")
        args.docx = False
    build_any = bool(args.docx)
    if not build_any:
        log("[*] --docx not given: download only")

    configure_docx(line_spacing=args.line_spacing, tafsir_font=args.tafsir_font, page_bg=args.page_bg,
                   ligatures=args.ligatures, verse_heuristic=args.verse_heuristic,
                   verse_color=args.verse_color)

    load_source_registry(args.sources_file)
    _common.LAYOUT = args.layout
    if args.html:
        cats = parse_categories_html(Path(args.html))
        if not any(cats.values()):
            p.error(f"no categories found in {args.html}")
        set_categories(cats)
        log(f"[*] categories read from {args.html}")

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

    surahs = parse_surah_spec(args.surah)
    if args.ayahs and len(surahs) != 1:
        p.error("--ayahs only works with a single --surah")
    if args.ayahs:
        parse_ayah_range(args.ayahs, surahs[0])   # validate early

    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    prefilter = not args.no_prefilter
    # cached sources.json (from an earlier run) lets us decide what is missing WITHOUT the network
    set_site_meta(load_json(out_dir / "sources.json", {}))   # group/category fallback for src_dir()
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

    # a slug the site does not have stops here — before any folder, file or ayah request
    validate_sources(get_session, args.base, out_dir, sources, args.allow_unknown)
    set_site_meta(load_json(out_dir / "sources.json", {}))
    if prefilter:   # the cached sources.json may have just been refreshed
        avail = Availability(load_json(out_dir / "sources.json", {}) or None)

    # ── 1. plan, from disk only: which (source, surah) really need the network? ──
    todo = []       # (src, surah, lo, hi, n_missing_ayahs)
    complete = 0
    no_material = 0
    for src in sources:
        for s in surahs:
            lo, hi = parse_ayah_range(args.ayahs, s)
            if not has_material(avail, src, s, prefilter, lo, hi):
                no_material += 1   # the source has nothing here: create NO files for it
                continue
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
    nm = f", {no_material} with no material in the source (skipped)" if no_material else ""
    if todo:
        log(f"[*] {total} (source, surah) pair(s): {complete} already on disk{nm}, {len(todo)} to "
            f"download (up to {sum(t[4] for t in todo)} ayah(s) not saved yet), {args.workers} workers")
    else:
        log(f"[*] {total} (source, surah) pair(s): everything is already on disk{nm} — nothing to download")

    # ── 2. network, only if something is missing ──
    ok = failed = 0
    skipped = complete
    if todo:
        if args.no_prefilter:
            atomic_write_json(out_dir / "sources.json",
                              merge_titles(load_json(out_dir / "sources.json", {}) or {}))
        else:
            meta = fetch_sources_meta(get_session(), args.base, out_dir)
            set_site_meta(meta)
            avail = Availability(meta)
            # sources.json is now fresh: drop surahs the source has no material for
            kept = [t for t in todo if has_material(avail, t[0], t[1], prefilter, t[2], t[3])]
            no_material += len(todo) - len(kept)
            todo = kept

    quran = None
    if args.with_quran and args.docx:
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
        report_incomplete(out_dir, sources, surahs, avail, prefilter)

    if not build_any:
        log("[*] raw data saved — nothing built (pass --docx to build Word files)")
    else:
        if args.docx and args.merge and not args.no_info:   # intros only go into full (merged) files
            need = ([] if (combine_order and args.combine_only) else list(sources)) + combine_order
            for src in dict.fromkeys(need):
                SOURCE_INTRO[src] = fetch_source_info(get_session(), args.base, out_dir, src, args.force)
            n_i = sum(1 for s_ in SOURCE_INTRO.values() if s_)
            log(f"[*] book intro: {n_i} of {len(SOURCE_INTRO)} source(s) have one")
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
