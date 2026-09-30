#!/usr/bin/env python3
"""
Scrape iPhone listings from https://www.refurbed.de into an Excel file.

Columns: variant, storage, price, seller_name, ts_data_pull

Respects polite rate limits. Use for personal / research purposes and follow
refurbed's Terms of Service and robots.txt.
"""

from __future__ import annotations

import argparse
import json
import re
import threading
import time
from datetime import datetime, timezone
from typing import Any
from urllib.parse import urljoin

import pandas as pd
import requests
from bs4 import BeautifulSoup

BASE_URL = "https://www.refurbed.de"
IPHONE_SITEMAP = f"{BASE_URL}/sitemap-marketplace/iphones/products.xml"
USER_AGENT = (
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
)


class RefurbedIphoneScraper:
    def __init__(self, delay: float = 0.6, timeout: int = 30) -> None:
        self.delay = delay
        self.timeout = timeout
        self.session = requests.Session()
        self.session.headers.update(
            {
                "User-Agent": USER_AGENT,
                "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
                "Accept-Language": "de-DE,de;q=0.9,en;q=0.8",
            }
        )
        self._last_request = 0.0
        self._rate_lock = threading.Lock()

    def _get(self, url: str) -> requests.Response:
        with self._rate_lock:
            elapsed = time.monotonic() - self._last_request
            if elapsed < self.delay:
                time.sleep(self.delay - elapsed)
            response = self.session.get(url, timeout=self.timeout)
            self._last_request = time.monotonic()
        response.raise_for_status()
        return response

    def get_model_urls(self) -> list[str]:
        xml = self._get(IPHONE_SITEMAP).text
        urls = sorted(
            set(re.findall(r"https://www\.refurbed\.de/p/[a-z0-9-]+/", xml))
        )
        return urls

    @staticmethod
    def _parse_ld_json_blocks(html: str) -> list[Any]:
        blocks: list[Any] = []
        for raw in re.findall(
            r'<script type="application/ld\+json">(.*?)</script>',
            html,
            flags=re.DOTALL,
        ):
            try:
                data = json.loads(raw)
            except json.JSONDecodeError:
                continue
            if isinstance(data, list):
                blocks.extend(data)
            else:
                blocks.append(data)
        return blocks

    def parse_variants_from_model_page(self, html: str, model_url: str) -> list[dict]:
        rows: list[dict] = []
        for item in self._parse_ld_json_blocks(html):
            if not isinstance(item, dict):
                continue
            if item.get("@type") != "ProductGroup":
                continue
            for variant in item.get("hasVariant") or []:
                if not isinstance(variant, dict):
                    continue
                offer = variant.get("offers") or {}
                if isinstance(offer, list):
                    offer = offer[0] if offer else {}
                price = offer.get("price")
                offer_url = offer.get("url") or ""
                if offer_url and offer_url.startswith("/"):
                    offer_url = urljoin(BASE_URL, offer_url)
                price_value = float(price) if price is not None else None
                # Skip unavailable / placeholder offers.
                if price_value is None or price_value <= 0 or not offer_url:
                    continue
                rows.append(
                    {
                        "variant": variant.get("name") or item.get("name"),
                        "storage": variant.get("size") or "",
                        "price": price_value,
                        "seller_name": "",
                        "offer_url": offer_url,
                        "model_url": model_url,
                        "color": variant.get("color") or "",
                        "sku": variant.get("sku"),
                    }
                )
        return rows

    def extract_seller_name(self, html: str) -> str:
        soup = BeautifulSoup(html, "lxml")
        node = soup.select_one('[data-test="merchant-name"]')
        if node:
            return node.get_text(" ", strip=True)

        # Fallback: GA / tracking payload sometimes includes merchant info nearby.
        match = re.search(
            r'data-test="product-page-merchant-link".*?</a>',
            html,
            flags=re.DOTALL,
        )
        if match:
            text = BeautifulSoup(match.group(0), "lxml").get_text(" ", strip=True)
            if text:
                return text
        return ""

    def enrich_with_seller(self, row: dict) -> dict:
        url = row.get("offer_url")
        if not url:
            return row
        try:
            html = self._get(url).text
            row["seller_name"] = self.extract_seller_name(html)
        except requests.RequestException as exc:
            row["seller_name"] = ""
            row["error"] = str(exc)
        return row

    def scrape(
        self,
        fetch_sellers: bool = True,
        limit_models: int | None = None,
    ) -> list[dict]:
        model_urls = self.get_model_urls()
        if limit_models is not None:
            model_urls = model_urls[:limit_models]

        print(f"Found {len(model_urls)} iPhone model pages", flush=True)
        rows: list[dict] = []
        for i, model_url in enumerate(model_urls, start=1):
            print(f"[{i}/{len(model_urls)}] {model_url}", flush=True)
            try:
                html = self._get(model_url).text
                model_rows = self.parse_variants_from_model_page(html, model_url)
                print(f"  -> {len(model_rows)} variants", flush=True)
                rows.extend(model_rows)
            except requests.RequestException as exc:
                print(f"  !! failed: {exc}", flush=True)

        if fetch_sellers and rows:
            print(f"Fetching seller names for {len(rows)} offers", flush=True)
            enriched: list[dict] = []
            for i, row in enumerate(rows, start=1):
                enriched.append(self.enrich_with_seller(dict(row)))
                if i % 25 == 0 or i == len(rows):
                    print(f"  sellers {i}/{len(rows)}", flush=True)
            rows = enriched

        ts = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")
        for row in rows:
            row["ts_data_pull"] = ts
        return rows


def to_excel(rows: list[dict], output_path: str) -> None:
    frame = pd.DataFrame(rows)
    columns = ["variant", "storage", "price", "seller_name", "ts_data_pull"]
    for col in columns:
        if col not in frame.columns:
            frame[col] = ""
    frame = frame[columns].fillna("")
    frame.to_excel(output_path, index=False)
    print(f"Wrote {len(frame)} rows -> {output_path}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Scrape refurbed.de iPhone offers into Excel"
    )
    parser.add_argument(
        "-o",
        "--output",
        default="refurbed_iphones.xlsx",
        help="Output Excel path (default: refurbed_iphones.xlsx)",
    )
    parser.add_argument(
        "--delay",
        type=float,
        default=0.6,
        help="Minimum seconds between HTTP requests (default: 0.6)",
    )
    parser.add_argument(
        "--skip-sellers",
        action="store_true",
        help="Skip seller page visits (faster; seller_name will be empty)",
    )
    parser.add_argument(
        "--limit-models",
        type=int,
        default=None,
        help="Only scrape the first N model pages (useful for testing)",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    scraper = RefurbedIphoneScraper(delay=args.delay)
    rows = scraper.scrape(
        fetch_sellers=not args.skip_sellers,
        limit_models=args.limit_models,
    )
    if not rows:
        raise SystemExit("No rows scraped. Site structure may have changed.")
    to_excel(rows, args.output)


if __name__ == "__main__":
    main()
