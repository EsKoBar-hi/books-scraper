# books-scraper

A simple Python web scraper for the sandbox book catalog [books.toscrape.com](http://books.toscrape.com).

It collects the following data from the first page of the catalog:
- **title** — book title
- **price** — price in GBP
- **in_stock** — stock status

The results are saved to a formatted Excel file (`books.xlsx`).

## Tech stack

- `requests` — for fetching the web page
- `BeautifulSoup` (bs4) — for parsing HTML
- `pandas` + `openpyxl` — for exporting data to Excel

## Installation

```bash
pip install requests beautifulsoup4 pandas openpyxl
```

## Usage

```bash
python parser.py
```

After running, a `books.xlsx` file will be created in the same folder, containing 20 books from the first page of the catalog.

## Example output

| title | price | in_stock |
|---|---|---|
| A Light in the Attic | £51.77 | In stock |
| Tipping the Velvet | £53.74 | In stock |
| Soumission | £50.10 | In stock |

## Possible improvements

- Scrape all catalog pages, not just the first one (pagination)
- Add book rating and product page URL
- Add a delay between requests to be polite to the server
