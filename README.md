# books-scraper

Simple Python web scrapers for the sandbox book catalog [books.toscrape.com](http://books.toscrape.com).

This repo contains two versions of the scraper, showing progression from a basic single-page scraper to a full multi-page one.

## Scripts

### `parser.py` — single page

Collects data from the **first page only** (20 books) and saves it to `books.xlsx`.

```bash
python parser.py
```

### `parser_all_pages.py` — full catalog

Follows the "next page" links automatically and collects data from the **entire catalog** (~1000 books, ~50 pages), with a short delay between requests to avoid overloading the server. Saves the result to `all_books.xlsx`.

```bash
python parser_all_pages.py
```

## Data collected

Both scripts collect the same fields for every book:
- **title** — book title
- **price** — price in GBP
- **in_stock** — stock status

Both output files are formatted Excel spreadsheets, with column widths auto-adjusted so text isn't cut off.

## Tech stack

- `requests` — for fetching web pages
- `BeautifulSoup` (bs4) — for parsing HTML
- `pandas` + `openpyxl` — for exporting data to Excel

## Installation

```bash
pip install requests beautifulsoup4 pandas openpyxl
```

## Example output

| title | price | in_stock |
|---|---|---|
| A Light in the Attic | £51.77 | In stock |
| Tipping the Velvet | £53.74 | In stock |
| Soumission | £50.10 | In stock |

## Possible improvements

- Add book rating and product page URL
- Add command-line arguments (e.g. `--url`, `--pages`) to make the scraper reusable for other sites
- Add a Flask/FastAPI wrapper to serve the data as an API
  
