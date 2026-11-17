# Tafseer Scraper

A Python command-line tool that downloads tafsir (Quran commentary) and related sources from [tafsir.app](https://tafsir.app) (الباحث القرآني) and exports them as **Markdown**, **Word (.docx)**, JSON and plain text, with proper Arabic right-to-left formatting.

## Features

- **Many sources, any surahs**: one source or several; one surah, a list or ranges (`1,23,44` or `1-3,53,112-114`) or the whole Quran (`all`).
- **One file per surah by default**, for every output format (Markdown, Word, text, JSON, and combined files). Add `--merge` to get one file per format holding all requested surahs instead.
- **Resumable**: each (source, surah) is saved as an append-only `.jsonl` file with a `.done` marker, so an interrupted run picks up where it stopped. Re-running the same command also retries any tasks that failed.
- **Parallel**: runs (source, surah) tasks on a thread pool (`--workers`, `--delay`).
- **Group-aware**: when the site returns one block covering several ayahs (for example 33-54), the scraper skips the ayahs already covered instead of requesting them again.
- **Skips empty ayahs**: uses the site's `sources.json` metadata to avoid ayahs a source has no material for (turn off with `--no_prefilter`).
- **Source groups**: add every source of a tab at once with `--group tafsir` or `--group uloom`.
- **Combine tafsirs**: merge several sources into one file, ordered per ayah: the verse, then each tafsir in the order you choose. Sources that are not downloaded yet are scraped automatically, and small typos in slugs are auto-corrected (for example `ibn-alqayim` becomes `ibn-alqayyim`).
- **Polished Word output** (single-source and combined files):
  - RTL layout with real Heading styles (usable in Word's Navigation Pane)
  - page numbers and a soft paper-colour background
  - Quran verses inside the text highlighted in a colour (`﴿…﴾`, plus vowelled quotes inside `( )`, `« »`, `[ ]`)
  - the `﴿ayah﴾` heading printed above each entry
  - footnotes
  - **poetry lines centred and in red**: verse lines (marked by the site with `؎`, with `∗∗∗` between the two halves) are shown centred in red, without the stray `؎` sign, and with a wide gap between the two halves
  - lines of `* * *` shown centred in red
  - the basmala on its own line
  - honorific ligatures (such as ﵇) spelled out when the font can't draw them
  - a download-date line in the document header
- **Markdown output** that works in any viewer (plain Markdown, no HTML). Verses are shown in **bold**, and poetry lines lose the stray `؎` and get a gap between the two halves (plain Markdown cannot centre or colour text; the Word output does).
- **Extra formats**: plain text and JSON, with the same per-surah / `--merge` rule as every other format.
- **Utilities**: list the known sources, show progress, and inspect unusual symbols in the saved data.

## Installation

Requires Python 3.10 or newer.

```bash
pip install curl_cffi python-docx
```

- `curl_cffi` is optional (it falls back to `requests`; install `requests` if you skip it).
- `python-docx` is only needed for `--docx`. Without it, DOCX output is skipped with a warning.

### Font

The Word output uses the **Kitab** font by default. Fonts are not embedded in the `.docx`, so Kitab must be installed on the machine that opens the file, otherwise Word substitutes another font. Use `--font "Name"` to choose a different one. Word accepts a single font name, so only the first name of a list such as `"Kitab, system-ui"` is used.

## Usage

```bash
# One surah, one source
python tafsir_scraper.py --src saadi --surah 53

# A range of ayahs inside one surah
python tafsir_scraper.py --src saadi --surah 53 --ayahs 33-54

# Several sources, several surahs
python tafsir_scraper.py --src saadi,muyassar --surah 1-3,53,112-114

# The whole Quran with one source (resumable)
python tafsir_scraper.py --src saadi --surah all

# Every source of the tafsir tab for a surah
python tafsir_scraper.py --group tafsir --surah 53

# Word output
python tafsir_scraper.py --src ibn-katheer --surah 53 --docx
python tafsir_scraper.py --src saadi --surah 1-3 --docx --font "Kitab" --docx_size 18

# Some surahs only (one file per surah in every format, the default)
python tafsir_scraper.py --src tadabbur-wa-amal --surah 1,23,44 --docx --txt

# The same surahs, but all in ONE file per format
python tafsir_scraper.py --src tadabbur-wa-amal --surah 1,23,44 --docx --txt --merge

# The whole Quran, one Word file and one Markdown file per surah
python tafsir_scraper.py --src tadabbur-wa-amal --surah all --docx

# Merge several tafsirs into one file (order = order given)
python tafsir_scraper.py --combine saadi,ibn-alqayyim,muyassar --surah 53 --docx
python tafsir_scraper.py --combine saadi,muyassar --surah 1-3 --docx --combine_only

# Example: Al-Mulk to An-Nas, one combined Word file per surah, five tafsirs in this order, no Markdown
python tafsir_scraper.py --combine almuyassar-ghareeb,mukhtasar,aysar-altafasir,wahidi-asbab,ibn-katheer --surah 67-114 --docx --combine_only --no-md

# Extras
python tafsir_scraper.py --list_sources                        # show known source slugs
python tafsir_scraper.py --src saadi --surah all --status      # progress only
python tafsir_scraper.py --src saadi --surah 53 --force        # re-scrape
python tafsir_scraper.py --src saadi --surah 53 --json --txt
python tafsir_scraper.py --src saadi --surah 53 --debug_marks  # list unusual symbols in the saved data
```

Run `python tafsir_scraper.py --help` for the full list of options.

### Options

**Selecting what to download**

| Option | Description |
| --- | --- |
| `--src` | Comma-separated source slugs (for example `saadi,muyassar`) |
| `--group` | Add every source of a tab: `tafsir` or `uloom` |
| `--surah` | `all`, a number, or a list/ranges such as `1,23,44` or `1-3,53` |
| `--ayahs` | Ayah range inside a single surah (for example `33-54`); only works with one `--surah` |
| `--combine` | Comma-separated sources, in order, merged into one file per ayah (verse, then each tafsir). Missing sources are scraped first |
| `--combine_only` | With `--combine`: skip the individual per-source outputs and build only the combined files |

At least one of `--src`, `--group` or `--combine` is required, together with `--surah` (except for `--list_sources`).

**Download behaviour**

| Option | Description |
| --- | --- |
| `--out_dir` | Output folder (default `./tafsir_output`) |
| `--workers` | Number of parallel (source, surah) tasks (default 4) |
| `--delay` | Pause in seconds after each request, per worker (default 0.3). Be polite to the site |
| `--force` | Re-scrape even if a surah is already marked done |
| `--no_prefilter` | Request every ayah even if `sources.json` says the source has no data for it |
| `--user-agent`, `--ua` | Custom User-Agent string for requests |

**Output formats and layout**

| Option | Description |
| --- | --- |
| `--md` / `--no-md` | Markdown output (on by default). `--no-md` also disables the combined Markdown files |
| `--docx` | Word output (needs `python-docx`) |
| `--txt` | Plain-text output |
| `--json` | JSON output (raw fields plus cleaned `text` and `footnotes`) |
| `--merge` | Put all requested surahs in **one file** per format, including `--combine`. Without it, every format gets one file per surah (the default) |

**Word (DOCX) appearance**

| Option | Description |
| --- | --- |
| `--font` | Word font (default `Kitab`); must be installed where the file is opened |
| `--docx_size` | Body size in pt (default 18) |
| `--line_spacing` | Line-spacing multiple (default 1.05; 1.0 = tight, 1.25 = airy) |
| `--page_bg` | Hex colour for the whole page background (default `F3EEE2`, soft paper); `none` for a white page |
| `--verse_color` | Hex colour for `﴿…﴾` quotations inside the tafsir text (default `0F6B9C`); `none` to disable. Used in DOCX (and in Markdown with `--md_verse_style html`) |
| `--verse_heuristic` / `--no-verse_heuristic` | Also colour vowelled quotes inside `( )`, `« »` and `[ ]` (on by default) |
| `--with_quran` / `--no-with_quran` | Print the `﴿ayah﴾` heading above each entry (on by default) |
| `--ligatures` | Honorific ligatures such as ﵇: `auto` (default) spells them out but keeps ﷺ, `expand` spells out all including ﷺ, `keep` leaves them as they are |

**Markdown appearance**

| Option | Description |
| --- | --- |
| `--md_verse_style` | How verses inside the tafsir are marked in Markdown: `bold` (default, works everywhere), `html` (coloured `<span>`, only where HTML renders) or `none` |

**Utilities**

| Option | Description |
| --- | --- |
| `--list_sources` | Print the known source slugs and exit |
| `--status` | Show progress per source and exit |
| `--debug_marks` | List unusual symbols found in the saved data (to see which quote marks the site uses) and exit |

## Output layout

File names use the Arabic surah name. With `--merge` and several surahs, the file name is `سورة <first> إلى سورة <last>` (only the surahs you asked for are inside), and for all 114 it is `القرآن كاملا`.

The rule is the same for every format (`md`, `txt`, `json`, `docx`):

```
tafsir_output/
├── quran.json                      # cached Quran text (used for the ﴿ayah﴾ headings)
├── sources.json                    # source metadata
├── <source>/
│   ├── raw/
│   │   ├── surah_053.jsonl         # raw records (source of truth)
│   │   └── surah_053.done          # completion marker
│   │
│   │   # default: one file per surah, one folder per format
│   ├── md/    سورة النجم.md
│   ├── docx/  سورة النجم.docx
│   ├── txt/   سورة النجم.txt
│   ├── json/  سورة النجم.json
│   │
│   │   # with --merge: one file per format, all surahs inside
│   └── <name>.md  <name>.docx  <name>.txt  <name>.json
└── combined/                       # with --combine (md and docx)
    ├── md/    سورة النجم.md         # default
    ├── docx/  سورة النجم.docx       # default
    └── <name>.md  <name>.docx      # with --merge
```

With `--no-md` no `.md` files are written (per source or combined).

## How it works

For each ayah the scraper requests `https://tafsir.app/get.php?src=<src>&s=<surah>&a=<ayah>&ver=1`, which returns the tafsir text (plus `ayahs_start` and `count` when the source comments on a group of ayahs in one block). The raw records are saved first as JSONL, and all other formats (Markdown, Word, JSON, TXT) are generated from them. That is why re-running with a different output option does not need to download anything again.

The text is cleaned before output: HTML tags and page markers are removed, footnotes become numbered `(١)` markers in the text with the notes listed after it, and the basmala and honorific ligatures are handled as described above.

## Notes

- This is an unofficial tool and is not affiliated with tafsir.app. Keep the request rate reasonable (the defaults are conservative) and respect the site's terms of use.
- Some sources (for example `wahidi-asbab`) only have material for certain ayahs; ayahs without material are skipped for that source.
- The tafsir texts belong to their respective authors, editors and publishers. The generated files are intended for personal study and research.