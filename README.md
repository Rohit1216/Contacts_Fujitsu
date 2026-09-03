# LinkedIn Team Deck Generator

Fills a branded PowerPoint "team card" template (org-chart-style slides with
circular photos + hyperlinked names + titles, plus an optional contact
table) from an Excel workbook — one sheet per team. Auto-paginates when a
team has more people than the template has card slots.

## How it works

- **`utils/excel_loader.py`** — reads the workbook. One sheet per team.
  Expected columns (any order, case-insensitive):
  `Name | Phone | Email | Designation | Location | PhotoUrl | LinkedinUrl`
- **`utils/ppt_generator.py`** — the engine:
  - `find_person_slots(slide)` auto-detects each photo+name+title "card" on
    a template slide purely by shape geometry — no manual coordinate mapping
    needed, so it adapts to a different template without code changes.
  - `fetch_circular_photo()` downloads each `PhotoUrl`, center-crops it to a
    square, masks it into a circle, and falls back to a colored initials
    avatar if the URL is missing or the download fails (with retries).
  - `paginate_team_sheet()` fills the template slide, then duplicates it as
    many times as needed for the rest of a sheet's people (title auto-updates
    to "(2/3)", "(3/3)", etc.), and deletes unused card slots on a partial
    final page.
  - `paginate_contact_table()` does the same for a table-style slide,
    growing/cloning table rows and paginating across slides as needed.
- **`generate.py`** — CLI entry point, useful for scripting/testing.
- **`app.py`** — Streamlit UI: upload template + workbook, map each sheet to
  a template slide, click Generate, download the result.

## 1. Local setup

```bash
git clone https://github.com/Rohit1216/ExcelPPTConvertorFuj.git
cd ExcelPPTConvertorFuj
python -m venv venv
source venv/bin/activate        # Windows: venv\Scripts\activate
pip install -r requirements.txt
```

## 2. Run it locally

**CLI** (fastest way to test a template + workbook combo):

```bash
python generate.py template.pptx workbook.xlsx output.pptx
```

Edit the `SHEET_SLIDE_MAP` dict at the top of `generate.py` so each sheet
name points at the right 0-based template slide index, e.g.:

```python
SHEET_SLIDE_MAP = {
    "Executive Management": 0,
    "IT": 1,
}
CONTACT_TABLE_SLIDE = 2
```

**Streamlit app** (no code editing — map sheets to slides in the UI):

```bash
streamlit run app.py
```

Opens at `http://localhost:8501`. Upload your template + workbook, choose
which slide each sheet fills, choose a contact-table slide, click
**Generate deck**, then **Download**.

## 3. Push to your GitHub repo

```bash
git add .
git commit -m "Team deck generator: geometry-based slot detection, autofit, pagination"
git push origin main
```

## 4. Deploy on Streamlit Community Cloud (optional)

1. Go to [share.streamlit.io](https://share.streamlit.io) → **New app**
2. Pick your repo/branch, set **Main file path** to `app.py`
3. Deploy — Streamlit Cloud installs `requirements.txt` automatically
4. Share the resulting `*.streamlit.app` URL

## Notes / things to know

- **Network access to photo URLs**: the app needs outbound HTTPS access to
  wherever `PhotoUrl` points (e.g. `media.licdn.com`). If you're running
  this inside a locked-down sandbox/CI environment, photo domains may be
  blocked and every photo will silently fall back to an initials avatar —
  it still runs, it just won't have real photos. Locally and on Streamlit
  Cloud this is a non-issue.
- **Template requirements**: each "card" on a team-grid slide must be a
  picture shape with two text boxes to its right (name above title) —
  that's what `find_person_slots` looks for. A background rectangle behind
  the card is optional but supported (it gets deleted along with the rest
  of the card on unused slots).
- **Long names/titles**: font size auto-shrinks (down to 7pt) if text would
  overflow its box — tune `min_pt` in `_autofit_shrink()` if you want a
  different floor.
- **Photo caching**: `fetch_circular_photo()` takes a shared `photo_cache`
  dict so the same LinkedIn URL (e.g. someone listed on two sheets) is only
  downloaded once per run.
- **Different template layout?** No code changes needed as long as each
  card is (picture + 2 textboxes), geometry detection adapts automatically.
  If your table slide's title also carries a "(x/y)" page counter, pass
  `title_prefix=` to `paginate_contact_table()` to keep it updating too.
