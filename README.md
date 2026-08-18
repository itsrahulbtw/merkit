# Flipkart FSN image utility

Reads an `.xlsx` file containing an `fsn` column, writes the Flipkart product URL into the `link` column, downloads the product's high-resolution gallery (hero) images, and records the result in the exported workbook.

```powershell
python -m pip install -r requirements.txt
python flipkart_fsn_scraper.py fsn.xlsx
```

The original spreadsheet is preserved. Each run creates:

```text
output/YYYY-MM-DD/
  fsn_scraped.xlsx
  images/<FSN>/hero_01.jpg
```

The exported workbook includes `scrape status`, `hero images downloaded`, `image folder`, and `scrape detail` columns. A status of `Scraped` means all gallery URLs discovered on that product page were downloaded successfully. Use a respectful delay when running larger sheets; the default is one second per product.
