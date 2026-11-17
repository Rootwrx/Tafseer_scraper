# Tafseer Scraper

A Python command-line tool that downloads tafsir (Quran commentary) and related sources from [tafsir.app](https://tafsir.app) (الباحث القرآني) and exports them as **Markdown**, **Word (.docx)**, JSON and plain text, with proper Arabic right-to-left formatting.

## Features

- **Many sources, any surah**: one source or several, one surah, a range (`1-3,53,112-114`) or the whole Quran (`all`).
- **Resumable**: each (source, surah) is saved as an append-only `.jsonl` file with a `.done` marker, so an interrupted run picks up where it stopped.
- **Parallel**: runs (source, surah) tasks on a thread pool.
- **Group-aware**: when the site returns one block covering several ayahs (for example 33-54), the scraper skips the ayahs already covered instead of requesting them again.
- **Skips empty ayahs**: uses the site's `sources.json` metadata to avoid ayahs a source has no material for.
- **Combine tafsirs**: merge several sources into one file, ordered per ayah: the verse, then each tafsir in the order you choose.
- **Polished Word output**:
  - RTL layout with real Heading styles (usable in Word's Navigation Pane)
  - page numbers and a soft paper-colour background
  - Quran verses inside the text highlighted in a colour
  - footnotes
  - lines of `* * *` shown centred in red
  - honorific ligatures (such as ﵇) spelled out when the font can't draw them
- **Markdown output** for any viewer, with optional coloured styling for viewers that render HTML.

## Installation

Requires Python 3.10 or newer.

```bash
pip install curl_cffi python-docx
```

- `curl_cffi` is optional (it falls back to `requests`; install `requests` if you skip it).
- `python-docx` is only needed for `--docx`.

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
python tafsir_scraper.py --src saadi --surah 1-3 --docx --docx_split --font "Kitab" --docx_size 18

# Merge several tafsirs into one file (order = order given)
python tafsir_scraper.py --combine saadi,ibn-alqayyim,muyassar --surah 53 --docx
python tafsir_scraper.py --combine saadi,muyassar --surah 1-3 --docx --docx_split --combine_only

# Extras
python tafsir_scraper.py --list_sources                        # show known source slugs
python tafsir_scraper.py --src saadi --surah all --status      # progress only
python tafsir_scraper.py --src saadi --surah 53 --force        # re-scrape
python tafsir_scraper.py --src saadi --surah 53 --json --single --txt
```

Run `python tafsir_scraper.py --help` for the full list of options.

### Main options

| Option | Description |
| --- | --- |
| `--src` | Comma-separated source slugs (for example `saadi,muyassar`) |
| `--group` | Add every source of a tab: `tafsir` or `uloom` |
| `--surah` | `all`, a number, or a list/ranges such as `1-3,53` |
| `--ayahs` | Ayah range inside a single surah (for example `33-54`) |
| `--out_dir` | Output folder (default `./tafsir_output`) |
| `--workers`, `--delay` | Parallel tasks and pause between requests (be polite to the site) |
| `--force` | Re-scrape even if already done |
| `--docx`, `--docx_split` | Word output; one file per surah with `--docx_split` |
| `--font`, `--docx_size` | Word font (default `Kitab`) and body size (default 18 pt) |
| `--page_bg`, `--verse_color`, `--line_spacing` | Word page colour, verse highlight colour and line spacing |
| `--combine`, `--combine_only` | Merge several tafsirs into one file |
| `--json`, `--single`, `--txt` | Extra output formats |
| `--ligatures` | `auto` (default), `expand` or `keep` for honorific ligatures |

## Output layout

```
tafsir_output/
├── quran.json                  # cached Quran text (used for the ﴿ayah﴾ headings)
├── sources.json                # source metadata
├── <source>/
│   ├── raw/
│   │   ├── surah_053.jsonl     # raw records (source of truth)
│   │   └── surah_053.done      # completion marker
│   ├── md/
│   │   └── سورة النجم.md        # per-surah Markdown
│   └── سورة النجم.docx          # with --docx
└── combined/                   # with --combine
    └── سورة النجم.md / .docx
```

## How it works

For each ayah the scraper requests `https://tafsir.app/get.php?src=<src>&s=<surah>&a=<ayah>&ver=1`, which returns the tafsir text. The raw records are saved first as JSONL, and all other formats (Markdown, Word, JSON, TXT) are generated from them. That is why re-running with a different output option does not need to download anything again.

## Notes

- This is an unofficial tool and is not affiliated with tafsir.app. Keep the request rate reasonable (the defaults are conservative) and respect the site's terms of use.
- The tafsir texts belong to their respective authors, editors and publishers. The generated files are intended for personal study and research.
