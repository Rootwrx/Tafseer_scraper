# Tafseer Scraper

A Python command-line tool that downloads tafsir (Quran commentary) and related sources from [tafsir.app](https://tafsir.app) (الباحث القرآني) and exports them as **Markdown**, **Word (.docx)**, JSON and plain text, with proper Arabic right-to-left formatting.

## Contents

- [Features](#features)
- [Installation](#installation)
- [Quick start](#quick-start)
- [Re-running, resuming and the network](#re-running-resuming-and-the-network)
- [Command guide](#command-guide)
- [Options](#options)
- [Rules between flags](#rules-between-flags)
- [Combined Word files](#combined-word-files)
- [Output layout](#output-layout)
- [Adding or renaming sources](#adding-or-renaming-sources)
- [How it works](#how-it-works)
- [Troubleshooting](#troubleshooting)
- [Notes](#notes)

## Features

- **Many sources, any surahs**: one source or several; one surah, a list or ranges (`1,23,44` or `1-3,53,112-114`) or the whole Quran (`all`).
- **Offline-first and resumable**: before any request, the script looks at what is already saved on disk and requests **only the ayahs that are missing**. If nothing is missing, it makes **zero network requests** and just rebuilds the output files. The run ends with the number of requests actually made. See [Re-running, resuming and the network](#re-running-resuming-and-the-network).
- **One file per surah by default**, for every output format (Markdown, Word, text, JSON, and combined files). Add `--merge` to get one file per format holding all requested surahs instead.
- **Choose your formats explicitly**: with no format flag you get Markdown + Word. As soon as you pass any of `--md`, `--docx`, `--json` or `--txt`, **only the formats you named** are built (so `--md` alone gives no Word file until you also add `--docx`). See [Choosing formats](#choosing-formats).
- **Parallel**: runs (source, surah) tasks on a thread pool (`--workers`, `--delay`), with automatic retry and back-off on 429 / 5xx / connection errors.
- **Group-aware**: when the site returns one block covering several ayahs (for example 33-54), the scraper skips the ayahs already covered instead of requesting them again.
- **Skips empty ayahs**: uses the site's `sources.json` metadata to avoid ayahs a source has no material for (turn off with `--no_prefilter`).
- **Source groups**: add every source of a tab at once with `--group tafsir` or `--group uloom`.
- **Arabic folder names**: output folders are named after the Arabic title of the source (for example `تفسير السعدي`), and folders created by older versions (English slugs, `combined/`) are renamed automatically.
- **Combine tafsirs**: merge several sources into one file, ordered per ayah: the verse, then each tafsir in the order you choose. Sources that are not downloaded yet are scraped automatically, and small typos in slugs are auto-corrected (for example `ibn-alqayim` becomes `ibn-alqayyim`).
- **Polished Word output** (single-source and combined files):
  - RTL layout with real Heading styles (usable in Word's Navigation Pane)
  - page numbers and a soft paper-colour background
  - Quran verses inside the text highlighted in a colour (`﴿…﴾`, plus vowelled quotes inside `( )`, `« »`, `[ ]`)
  - the `﴿ayah﴾` heading printed above each entry
  - footnotes, with `(١)` markers in the text
  - **poetry lines centred and in red**: verse lines (marked by the site with `؎`, with `∗∗∗` between the two halves) are shown centred in red, without the stray `؎` sign, and with a wide gap between the two halves
  - lines of `* * *` shown centred in red
  - the basmala on its own line
  - honorific ligatures (such as ﵇) spelled out when the font can't draw them
  - a source and download-date line in the document header
  - a separate font for tafsir names and text (`--tafsir_font`), while Quran verses stay in `--font`
  - **collapsible sections in combined files**: in Word, collapsing a verse hides all its tafsirs, and collapsing a tafsir name hides that tafsir's text and footnotes (see [Combined Word files](#combined-word-files))
  - optional **bold tafsir text** and adjustable **tafsir name size** in combined files
  - no stray italics: Arabic text is forced to upright in every heading level
- **Markdown output** that works in any viewer (plain Markdown, no HTML by default). Verses are shown in **bold**, and poetry lines lose the stray `؎` and get a gap between the two halves (plain Markdown cannot centre or colour text; the Word output does).
- **Extra formats**: plain text and JSON, with the same per-surah / `--merge` rule as every other format.
- **Utilities**: list the known sources, show progress (from disk, no network), and inspect unusual symbols in the saved data.

## Installation

Requires Python 3.10 or newer.

```bash
pip install curl_cffi python-docx
```

- `curl_cffi` is optional (it falls back to `requests`; install `requests` if you skip it).
- `python-docx` is only needed for Word output. Without it, the script prints one warning, disables DOCX and still builds the other formats.

### Font

The Word output uses the **Kitab** font by default for headings and Quran verses, and **KFGQPC Uthman Taha Naskh** for tafsir names and text. Fonts are not embedded in the `.docx`, so they must be installed on the machine that opens the file, otherwise Word substitutes another font. Use `--font "Name"` and `--tafsir_font "Name"` to choose different ones. Word accepts a single font name, so only the first name of a list such as `"Kitab, system-ui"` is used.

## Quick start

```bash
# One surah, one source (no format flag → Markdown + Word, one file each)
python tafsir_scraper.py --src saadi --surah 53

# Find source slugs
python tafsir_scraper.py --list_sources
```

Run the same command again at any time: it only downloads what is missing.

## Re-running, resuming and the network

The raw data is the source of truth. Every (source, surah) is saved as `raw/surah_NNN.jsonl`, and a `surah_NNN.done` marker is written once the whole surah is complete. Every other format (Markdown, Word, JSON, TXT) is generated from that raw data.

Each run works in two steps:

1. **Plan, from disk only.** For every (source, surah) the script checks the `.jsonl` / `.done` files and the cached `sources.json`, and works out which ayahs are still missing.
2. **Download, only if something is missing.** The network is used only for the missing ayahs. If nothing is missing, no session is even opened.

| Situation | Network requests |
| --- | --- |
| Fresh surah | `sources.json`, `quran.json` (cached afterwards), then one per ayah (or per grouped block) |
| Same command run again | **0** |
| `.done` marker deleted but all data saved | **0** (the marker is repaired) |
| One ayah missing | `sources.json` + that ayah |
| Added a new source to the command | Only that source's missing ayahs |
| Changed only output options (`--json`, `--merge`, `--font` …) | **0** |
| `--force` | Everything in the requested range again |

Notes:

- The only request a fully downloaded run can still make is `quran.json`, **once**, if it was never cached (used for the `﴿ayah﴾` headings). Use `--no-with_quran` to avoid it.
- `sources.json` is refreshed once per run that needs the network, and the cached copy is used to plan offline.
- `Ctrl+C` is safe: progress is saved and the next run resumes.
- If some tasks fail, the run ends with a warning; run the same command again to retry only those.
- Before building, the script warns about surahs that are not marked complete (the output may be partial). The warning is skipped when you use `--ayahs`.
- The final line shows the totals, for example: `downloaded=1 already-on-disk=1 failed=0 | network requests: 0 (everything was already on disk)`.
- `--status` shows what is saved and what still needs downloading, without touching the network.

## Command guide

### One source, one surah

```bash
python tafsir_scraper.py --src saadi --surah 53
```

One file per format, in a subfolder per format:

```
<out_dir>/تفسير السعدي/md/سورة النجم.md
<out_dir>/تفسير السعدي/docx/سورة النجم.docx
```

With one surah, `--merge` only moves the file out of the `md/` and `docx/` subfolders into `<out_dir>/تفسير السعدي/`; the content and name are the same.

### An ayah range

```bash
python tafsir_scraper.py --src saadi --surah 53 --ayahs 33-54
```

Only that range is downloaded (a grouped block that sticks out of the range is kept whole). No `.done` marker is written, because the surah is not complete. The output files are built from **everything saved** for that surah, not only the range. `--ayahs` works with a single surah only.

### Several sources, all surahs, no `--merge`

```bash
python tafsir_scraper.py --src saadi,ibn-katheer --surah all
```

The two tafsirs stay separate. For each source you get one file per surah, per format:

```
<out_dir>/تفسير السعدي/raw/surah_001.jsonl        (+ .done)
<out_dir>/تفسير السعدي/md/سورة الفاتحة.md          … 114 files
<out_dir>/تفسير السعدي/docx/سورة الفاتحة.docx       … 114 files
<out_dir>/تفسير ابن كثير/raw|md|docx/…
```

That is up to 2 × 114 × 2 = 456 output files plus the raw data. To build something else, name the formats (for example `--docx` for Word only, or `--md --json`).

### Several sources, all surahs, with `--merge`

```bash
python tafsir_scraper.py --src saadi,ibn-katheer --surah all --merge
```

Each source gets **its own single file per format**; the sources are never mixed:

```
<out_dir>/تفسير السعدي/القرآن كاملا.docx
<out_dir>/تفسير السعدي/القرآن كاملا.md
<out_dir>/تفسير ابن كثير/القرآن كاملا.docx
<out_dir>/تفسير ابن كثير/القرآن كاملا.md
```

- Each file holds all surahs in order. In Word every surah starts on a new page with its own heading.
- The files sit directly in the source folder (no `md/` or `docx/` subfolder). `--json` and `--txt` are merged the same way.
- The raw per-surah data is untouched.
- The file name reflects what the file really contains: `القرآن كاملا` for all 114, `سورة النجم إلى سورة الرحمن` for a contiguous range, and `سور 1-3،53،112-114` when the surahs have gaps (for example when a source has no data for some surahs). It never claims a range that has holes.
- A full-Quran Word file of a large tafsir can be very big and slow to open; use the per-surah files if that is a problem.

### Merge several tafsirs into one file (`--combine`)

`--merge` never mixes sources. To put several tafsirs in the **same** file, use `--combine`:

```bash
# one combined file per surah: verse, then saadi, then ibn-alqayyim, then muyassar
python tafsir_scraper.py --combine saadi,ibn-alqayyim,muyassar --surah 53

# only the combined files, no individual per-source files
python tafsir_scraper.py --combine saadi,muyassar --surah 1-3 --combine_only

# everything in ONE combined file
python tafsir_scraper.py --combine saadi,ibn-katheer --surah all --merge --combine_only
```

Output goes to one folder per combination (and order), so different combinations never overwrite each other:

```
<out_dir>/جمع_التفاسير/تفسير السعدي_تفسير ابن كثير/docx/سورة النجم.docx     (default)
<out_dir>/جمع_التفاسير/تفسير السعدي_تفسير ابن كثير/القرآن كاملا.docx       (with --merge)
```

Details:

- Sources that are not downloaded yet are scraped first; the ones already saved cost nothing.
- Without `--combine_only`, the individual per-source files are built too.
- Each ayah entry starts at every ayah where at least one source begins a block. If a source has a longer block that started earlier (for example one block for ayahs 33-54), the later entry shows an italic note pointing to it instead of repeating the text, and a tafsir heading shows its ayah range when it differs from the entry's.
- Sources with no text for a surah are left out.
- Combined output is Markdown and Word only; `--json` / `--txt` apply to the individual per-source files. Name the formats you want: `--combine a,b --surah 53 --docx` builds only the combined Word file.

### Source groups

```bash
python tafsir_scraper.py --group tafsir --surah 53     # every source of the "تفسير" tab
python tafsir_scraper.py --group uloom --surah 53      # every source of the "علوم" tab
python tafsir_scraper.py --group tafsir --src ibn-uthaymeen --surah 53   # groups and --src add up
```

### Choosing formats

```bash
python tafsir_scraper.py --src saadi --surah 53                  # no format flag → Markdown + Word
python tafsir_scraper.py --src saadi --surah 53 --md             # Markdown ONLY
python tafsir_scraper.py --src saadi --surah 53 --docx           # Word ONLY
python tafsir_scraper.py --src saadi --surah 53 --md --docx      # both
python tafsir_scraper.py --src saadi --surah 53 --json           # JSON ONLY
python tafsir_scraper.py --src saadi --surah 53 --md --json --txt    # exactly these three
python tafsir_scraper.py --src saadi --surah 1-9 --docx --merge  # one merged Word file
python tafsir_scraper.py --src saadi --surah all --no-md --no-docx   # download only
```

The rule: **no format flag → Markdown + Word. Any of `--md`, `--docx`, `--json`, `--txt` → only the formats you name.** `--no-md` / `--no-docx` only remove a format from the default pair (`--no-md` alone gives Word only).

### Rebuilding without downloading

There is no separate "build only" mode (and no `--build_only` flag): the script already does this by itself. Run the same command with different output options and it rebuilds from the saved raw data, with zero requests:

```bash
python tafsir_scraper.py --src saadi --surah all                 # first run: downloads
python tafsir_scraper.py --src saadi --surah all --merge --json  # second run: no network, JSON only
```

### Word appearance

```bash
python tafsir_scraper.py --src saadi --surah 1-3 --font "Kitab" --docx_size 18
python tafsir_scraper.py --src saadi --surah 53 --page_bg none --line_spacing 1.0

# combined Word file with bigger tafsir names, bold text and tighter lines
python tafsir_scraper.py --combine saadi,muyassar --surah 53 --tafsir_title_size 22 --tafsir_bold --line_spacing 0.85

# Al-Mulk to An-Nas, one combined Word file per surah, five tafsirs in this order, no Markdown
python tafsir_scraper.py --combine almuyassar-ghareeb,mukhtasar,aysar-altafasir,wahidi-asbab,ibn-katheer --surah 67-114 --combine_only --docx
```

### Progress, force and utilities

```bash
python tafsir_scraper.py --src saadi --surah all --status        # progress from disk, no network
python tafsir_scraper.py --src saadi --surah 53 --force          # re-download the whole surah
python tafsir_scraper.py --src saadi --surah 53 --ayahs 40-45 --force   # re-download only that range
python tafsir_scraper.py --src saadi --surah 53 --debug_marks    # list unusual symbols in the saved data
python tafsir_scraper.py --list_sources                          # known source slugs
```

`--status` prints one line per source, for example `saadi  done 114/114  partial 0  need download 0`.

Run `python tafsir_scraper.py --help` for the full list of options.

## Options

**Selecting what to download**

| Option | Description |
| --- | --- |
| `--src` | Comma-separated source slugs (for example `saadi,muyassar`) |
| `--group` | Add every source of a tab: `tafsir` or `uloom` |
| `--surah` | `all`, a number, or a list/ranges such as `1,23,44` or `1-3,53` |
| `--ayahs` | Ayah range inside a single surah (for example `33-54`); only works with one `--surah` |
| `--combine` | Comma-separated sources, in order, merged into one file per ayah (verse, then each tafsir). Missing sources are scraped first |
| `--combine_only` | Requires `--combine`: build only the combined files and skip the individual per-source outputs |

At least one of `--src`, `--group` or `--combine` is required, together with `--surah` (except for `--list_sources`).

**Download behaviour**

| Option | Description |
| --- | --- |
| `--out_dir` | Output folder (default `./tafsir_output`) |
| `--workers` | Number of parallel (source, surah) tasks (default 4) |
| `--delay` | Pause in seconds after each request, per worker (default 0.3). Be polite to the site |
| `--force` | Re-download even if already saved. With `--ayahs`, only that range. Without it, saved data is never requested again |
| `--no_prefilter` | Request every ayah even if `sources.json` says the source has no data for it |
| `--user-agent`, `--ua` | Custom User-Agent string for requests |

**Output formats and layout**

| Option | Description |
| --- | --- |
| `--md` / `--no-md` | Markdown output. Built by default only when **no** format flag is given; if you pass any of `--md/--docx/--json/--txt`, only the named ones are built. `--no-md` removes it from the default pair (also for `--combine`) |
| `--docx` / `--no-docx` | Word output (needs `python-docx`). Same rule as `--md` |
| `--txt` | Plain-text output (when given, md/docx are built only if you also name them) |
| `--json` | JSON output: raw fields plus cleaned `text` and `footnotes` (when given, md/docx are built only if you also name them) |
| `--merge` | Put all requested surahs in **one file** per format (and per source), including `--combine`. Without it, every format gets one file per surah (the default) |

**Word (DOCX) appearance**

| Option | Description |
| --- | --- |
| `--font` | Font for headings and Quran verses (default `Kitab`); must be installed where the file is opened |
| `--tafsir_font` | Font for tafsir names and tafsir text (default `KFGQPC Uthman Taha Naskh`). Quran verses stay in `--font` |
| `--docx_size` | Body size in pt (default 16). Headings, verses and footnotes scale from it |
| `--line_spacing` | Line-spacing multiple (default 0.9; 1.0 = the font's natural line height, 1.25 = airy). Also applied to the Title and Heading 1-4 styles. If Arabic marks (tashkeel) look clipped, raise it to 1.0 |
| `--tafsir_title_size` | `--combine` only: size in pt of the tafsir names, both the heading above each tafsir and the list on the cover page (default 20) |
| `--tafsir_bold` / `--no-tafsir_bold` | `--combine` only: bold tafsir body text in Word files (**off** by default). Poetry lines, the basmala, separators and footnotes keep their own styling |
| `--page_bg` | Hex colour for the whole page background (default `F3EEE2`, soft paper); `none` for a white page |
| `--verse_color` | Hex colour for `﴿…﴾` quotations inside the tafsir text (default `0F6B9C`); `none` to disable. Used in DOCX, and in Markdown only with `--md_verse_style html` |
| `--verse_heuristic` / `--no-verse_heuristic` | Also colour vowelled quotes inside `( )`, `« »` and `[ ]` (on by default) |
| `--with_quran` / `--no-with_quran` | Print the `﴿ayah﴾` heading above each entry (on by default). Turning it off also avoids downloading `quran.json` |
| `--ligatures` | Honorific ligatures such as ﵇: `auto` (default) spells them out but keeps ﷺ, `expand` spells out all including ﷺ, `keep` leaves them as they are |

**Markdown appearance**

| Option | Description |
| --- | --- |
| `--md_verse_style` | How verses inside the tafsir are marked in Markdown: `bold` (default, works everywhere), `html` (coloured `<span>`, only where HTML renders) or `none` |

**Utilities**

| Option | Description |
| --- | --- |
| `--list_sources` | Print the known source slugs (tafsir, uloom and extra) and exit |
| `--status` | Show progress per source from disk (done / partial / need download) and exit. No network |
| `--debug_marks` | List unusual symbols found in the saved data (to see which quote marks the site uses) and exit |
| `--sources_file` | Extra JSON file of Arabic source titles (see [Adding or renaming sources](#adding-or-renaming-sources)) |

## Rules between flags

The script checks flag combinations and tells you when something has no effect:

| Combination | Result |
| --- | --- |
| `--combine_only` without `--combine` | Error |
| `--ayahs` with more than one surah | Error (`--ayahs` needs a single `--surah`) |
| `--json` / `--txt` with `--combine` | Warning: only the individual per-source files get them; with `--combine_only` they are ignored |
| `--combine` without `--md` / `--docx` (for example only `--json`) | Warning: combined files are md/docx only, so none are built |
| `--tafsir_title_size` / `--tafsir_bold` without `--combine` | Warning: they only affect combined Word files |
| Word output requested but `python-docx` missing | One warning, DOCX is disabled, other formats are built |
| No format enabled (for example `--no-md --no-docx`, or `--docx` without `python-docx`) | Warning: nothing is built, the script only downloads |
| `--build_only` | Removed. Re-running the same command already rebuilds from disk without network |

## Combined Word files

In a `--combine` Word file the outline is:

| Level | Content | Collapsing it hides |
| --- | --- | --- |
| Heading 1 | Surah name | everything in the surah |
| Heading 2 | `الآية N` (ayah label) | the verse and all its tafsirs |
| Verse paragraph (outline level 3) | `﴿…﴾` | all tafsirs of that ayah |
| Heading 4 | Tafsir name, in the colour of its source | that tafsir's text and footnotes |

In Word, hover over a heading or the verse line to see the collapse arrow; right-click and choose **Expand/Collapse > Collapse All Headings** to fold everything at once. Because the verse has an outline level, it also appears in the Navigation Pane. This applies to `--combine` files only; single-source Word files keep the plain Heading 1-3 structure.

## Output layout

Folders are named with the Arabic title of the source; the technical subfolders (`raw`, `md`, `docx`, `txt`, `json`) stay in English. File names use the Arabic surah name.

The rule is the same for every format (`md`, `txt`, `json`, `docx`):

```
tafsir_output/
├── quran.json                      # cached Quran text (used for the ﴿ayah﴾ headings)
├── sources.json                    # source metadata (cached; used to plan offline)
├── <عنوان المصدر>/                  # one folder per source, e.g. تفسير السعدي
│   ├── raw/
│   │   ├── surah_053.jsonl         # raw records (source of truth)
│   │   └── surah_053.done          # completion marker
│   │
│   │   # default: one file per surah, one folder per format
│   ├── md/    سورة النجم.md
│   ├── docx/  سورة النجم.docx
│   ├── txt/   سورة النجم.txt        # with --txt
│   ├── json/  سورة النجم.json       # with --json
│   │
│   │   # with --merge: one file per format, all surahs inside
│   └── <name>.md  <name>.docx  <name>.txt  <name>.json
└── جمع_التفاسير/                    # with --combine (md and docx only)
    └── <مصدر١>_<مصدر٢>/             # one folder per combination and order
        ├── md/    سورة النجم.md      # default
        ├── docx/  سورة النجم.docx    # default
        └── <name>.md  <name>.docx   # with --merge
```

`<name>` in a merged file is built from the surahs it really contains:

| Surahs in the file | File name |
| --- | --- |
| all 114 | `القرآن كاملا` |
| one surah | `سورة النجم` |
| a contiguous range | `سورة النجم إلى سورة الرحمن` |
| anything else | `سور 1-3،53،112-114` |

Markdown files are only written when `--md` is given or no format flag is used at all (same for Word); `--no-md` always suppresses them, per source and combined. Folders created by older versions (English-slug source folders, `combined/`, Arabic-named format subfolders) are renamed automatically the first time you run the new version.

## Adding or renaming sources

The script ships with built-in lists of sources (`تفسير` and `علوم` tabs, plus a few non-text extras that must be passed explicitly with `--src`). To add a source or change an Arabic title without editing the code, create a `sources_ar.json` next to the script (or in the working directory, or pass `--sources_file path.json`):

```json
{
  "my-new-slug": { "title": "اسم المصدر", "group": "tafsir", "enabled": true },
  "saadi": "تفسير السعدي"
}
```

- A value can be a plain title string or an object with `title`, `group` (`tafsir` or `uloom`) and `enabled`.
- Sources with a `group` are included in `--group` runs unless `enabled` is `false`.
- Slugs that are not in any list are still tried, with a warning.

## How it works

For each ayah the scraper requests `https://tafsir.app/get.php?src=<src>&s=<surah>&a=<ayah>&ver=1`, which returns the tafsir text (plus `ayahs_start` and `count` when the source comments on a group of ayahs in one block). The raw records are saved first as JSONL, and all other formats (Markdown, Word, JSON, TXT) are generated from them. That is why re-running with a different output option does not need to download anything again.

Two small files are also fetched when needed: `sources/sources.json` (which ayahs each source covers, used to skip empty ayahs) and `sources/quran.json` (the Quran text for the `﴿ayah﴾` headings). Both are cached in the output folder.

The text is cleaned before output: HTML tags and page markers are removed, footnotes become numbered `(١)` markers in the text with the notes listed after it, and the basmala and honorific ligatures are handled as described above.

## Troubleshooting

- **`HTTP 403 — blocked`**: lower `--workers` and raise `--delay`, then run the same command again; it resumes where it stopped.
- **Some surahs show as not complete**: a download failed or was interrupted. Run the same command again to fetch only what is missing.
- **A source produced no files**: it has no material for the requested surahs (for example `wahidi-asbab` covers only some ayahs), or nothing was downloaded yet.
- **Boxes or wrong fonts in Word**: install the fonts (`Kitab`, `KFGQPC Uthman Taha Naskh`) or choose installed ones with `--font` / `--tafsir_font`. For honorific symbols that still show as boxes, try `--ligatures expand`.
- **Clipped Arabic marks (tashkeel)**: raise `--line_spacing` to 1.0 or more.
- **Wrong quote marks or odd symbols in the text**: run `--debug_marks` to see which characters the site uses.
- **Want a clean re-download**: use `--force` (add `--ayahs` to limit it to a range).

## Notes

- This is an unofficial tool and is not affiliated with tafsir.app. Keep the request rate reasonable (the defaults are conservative) and respect the site's terms of use.
- Some sources (for example `wahidi-asbab`) only have material for certain ayahs; ayahs without material are skipped for that source.
- The tafsir texts belong to their respective authors, editors and publishers. The generated files are intended for personal study and research.