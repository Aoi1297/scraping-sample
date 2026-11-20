"""
Webサイトの商品一覧を集めて、CSV と Excel に書き出すスクレイピングツールの見本。

対象: https://books.toscrape.com/ (スクレイピング練習用に公開されているサイト)

使い方:
  python scrape_books.py                  # 全ページ(50ページ・1000件)
  python scrape_books.py --pages 3        # 最初の3ページだけ
  python scrape_books.py --detail         # 詳細ページも開いてカテゴリ・在庫数・説明も取る
  python scrape_books.py --out result     # 出力ファイル名(result.csv / result.xlsx)

工夫している点:
  - アクセスの間隔をあける(初期値 1秒)/ 失敗したら待ってやり直す
  - CSV は Excel で文字化けしない形式(UTF-8 BOM付き)
  - Excel は見出しの固定・フィルター・列幅の調整つき
"""

import argparse
import csv
import logging
import re
import time
from pathlib import Path
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

BASE = "https://books.toscrape.com/catalogue/"
RATING = {"One": 1, "Two": 2, "Three": 3, "Four": 4, "Five": 5}
HEADERS = {"User-Agent": "Mozilla/5.0 (scraping-sample; +practice site)"}

log = logging.getLogger("scraper")


class Fetcher:
    """間隔をあけて取得し、失敗したら数回やり直す"""

    def __init__(self, delay: float, retries: int = 3):
        self.delay = delay
        self.retries = retries
        self.session = requests.Session()
        self.session.headers.update(HEADERS)
        self._last = 0.0

    def get(self, url: str) -> BeautifulSoup:
        for attempt in range(1, self.retries + 1):
            wait = self.delay - (time.time() - self._last)
            if wait > 0:
                time.sleep(wait)
            self._last = time.time()
            try:
                r = self.session.get(url, timeout=20)
                r.raise_for_status()
                r.encoding = "utf-8"
                return BeautifulSoup(r.text, "html.parser")
            except requests.RequestException as e:
                log.warning("取得失敗 (%d/%d) %s: %s", attempt, self.retries, url, e)
                time.sleep(2 * attempt)
        raise RuntimeError(f"取得できませんでした: {url}")


def parse_list(soup: BeautifulSoup, page_url: str) -> list[dict]:
    rows = []
    for item in soup.select("article.product_pod"):
        a = item.select_one("h3 a")
        stars = item.select_one("p.star-rating")["class"]
        price = item.select_one("p.price_color").get_text(strip=True)
        rows.append({
            "タイトル": a["title"],
            "価格(£)": float(re.sub(r"[^\d.]", "", price)),
            "評価(星)": next((RATING[c] for c in stars if c in RATING), None),
            "在庫": item.select_one("p.availability").get_text(strip=True),
            "URL": urljoin(page_url, a["href"]),
        })
    return rows


def parse_detail(soup: BeautifulSoup) -> dict:
    table = {th.get_text(strip=True): td.get_text(strip=True)
             for th, td in zip(soup.select("table th"), soup.select("table td"))}
    stock = re.search(r"\d+", table.get("Availability", ""))
    desc = soup.select_one("#product_description + p")
    return {
        "カテゴリ": soup.select("ul.breadcrumb li a")[-1].get_text(strip=True),
        "在庫数": int(stock.group()) if stock else None,
        "UPC": table.get("UPC", ""),
        "説明": desc.get_text(strip=True) if desc else "",
    }


def scrape(pages: int | None, detail: bool, delay: float) -> list[dict]:
    f = Fetcher(delay)
    rows, n = [], 1
    while pages is None or n <= pages:
        url = urljoin(BASE, f"page-{n}.html")
        soup = f.get(url)
        got = parse_list(soup, url)
        if detail:
            for r in got:
                r.update(parse_detail(f.get(r["URL"])))
        rows += got
        log.info("ページ %d: %d件 (合計 %d件)", n, len(got), len(rows))
        if not soup.select_one("li.next a"):
            break
        n += 1
    return rows


def save_csv(rows: list[dict], path: Path):
    with path.open("w", encoding="utf-8-sig", newline="") as fp:
        w = csv.DictWriter(fp, fieldnames=list(rows[0]))
        w.writeheader()
        w.writerows(rows)


def save_xlsx(rows: list[dict], path: Path):
    wb = Workbook()
    ws = wb.active
    ws.title = "商品一覧"
    cols = list(rows[0])
    ws.append(cols)
    for r in rows:
        ws.append([r[c] for c in cols])
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="305496")
        cell.alignment = Alignment(horizontal="center")
    widths = {"タイトル": 50, "URL": 60, "説明": 80}
    for i, c in enumerate(cols, start=1):
        ws.column_dimensions[ws.cell(1, i).column_letter].width = widths.get(c, 14)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(path)


def main():
    ap = argparse.ArgumentParser(description="商品一覧のスクレイピング見本")
    ap.add_argument("--pages", type=int, help="取得するページ数(省略で全ページ)")
    ap.add_argument("--detail", action="store_true", help="詳細ページも取得する")
    ap.add_argument("--delay", type=float, default=1.0, help="アクセス間隔(秒)")
    ap.add_argument("--out", default="books", help="出力ファイル名(拡張子なし)")
    args = ap.parse_args()

    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(message)s", datefmt="%H:%M:%S")
    rows = scrape(args.pages, args.detail, args.delay)
    if not rows:
        log.error("データが取れませんでした")
        return
    out = Path(args.out)
    save_csv(rows, out.with_suffix(".csv"))
    save_xlsx(rows, out.with_suffix(".xlsx"))
    log.info("保存しました: %s / %s (%d件)", out.with_suffix(".csv"), out.with_suffix(".xlsx"), len(rows))


if __name__ == "__main__":
    main()
