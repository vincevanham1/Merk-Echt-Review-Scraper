"""Utility for exporting Klantenvertellen reviews to an Excel file."""

import argparse
import json
from dataclasses import dataclass
from typing import Dict, Iterable, List, Optional

import pandas as pd
import requests
from bs4 import BeautifulSoup, Tag


@dataclass
class Review:
    reviewer: str
    score: str
    comment: str

    def to_dict(self) -> Dict[str, str]:
        return {
            "Reviewer Name": self.reviewer or "Anonymous",
            "Score": self.score or "N/A",
            "Comments": self.comment or "No comment",
        }


def _default_headers() -> Dict[str, str]:
    return {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
            "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
        ),
        "Accept-Language": "en-US,en;q=0.9,nl;q=0.8",
    }


def _safe_json_loads(payload: str) -> Optional[object]:
    try:
        return json.loads(payload)
    except json.JSONDecodeError:
        return None


def _extract_reviews_from_json(obj: object) -> List[Review]:
    reviews: List[Review] = []

    if isinstance(obj, dict):
        if obj.get("@type", "").lower() == "review":
            author = obj.get("author")
            rating = obj.get("reviewRating", {}).get("ratingValue") if isinstance(
                obj.get("reviewRating"), dict
            ) else obj.get("reviewRating")
            description = obj.get("description") or obj.get("reviewBody")
            if isinstance(author, dict):
                author = author.get("name")
            reviews.append(
                Review(
                    reviewer=str(author or "Anonymous").strip(),
                    score=str(rating or "N/A").strip(),
                    comment=str(description or "No comment").strip(),
                )
            )
        for value in obj.values():
            reviews.extend(_extract_reviews_from_json(value))
    elif isinstance(obj, list):
        for item in obj:
            reviews.extend(_extract_reviews_from_json(item))

    return reviews


def _parse_ld_json(soup: BeautifulSoup) -> List[Review]:
    collected: List[Review] = []
    for script in soup.find_all("script", type="application/ld+json"):
        payload = script.string or script.get_text()
        if not payload:
            continue
        data = _safe_json_loads(payload)
        if data:
            collected.extend(_extract_reviews_from_json(data))
    return collected


def _first_text(soup: BeautifulSoup, selectors: Iterable[str]) -> Optional[str]:
    for selector in selectors:
        node = soup.select_one(selector)
        if node and node.get_text(strip=True):
            return node.get_text(strip=True)
    return None


def _parse_dom_reviews(soup: BeautifulSoup) -> List[Review]:
    containers = soup.select(
        '[itemtype*="Schema.org/Review"], [itemtype*="schema.org/Review"], .review, .review-item, article'
    )
    parsed: List[Review] = []

    for container in containers:
        if not isinstance(container, Tag):
            continue
        name = _first_text(
            container,
            [
                "[itemprop=author]",
                "[itemprop='name']",
                ".reviewer-name",
                ".author",
                "header strong",
            ],
        )
        score = _first_text(
            container,
            [
                "[itemprop='ratingValue']",
                "[itemprop='reviewRating']",
                ".rating",
                ".score",
                "[class*='star']",
            ],
        )
        comment = _first_text(
            container,
            [
                "[itemprop='description']",
                "[itemprop='reviewBody']",
                ".review-text",
                ".comment",
                "[class*='content']",
            ],
        )

        if any([name, score, comment]):
            parsed.append(
                Review(
                    reviewer=(name or "Anonymous"),
                    score=(score or "N/A"),
                    comment=(comment or "No comment"),
                )
            )
    return parsed


def _fetch_page(session: requests.Session, url: str, page: int) -> Optional[str]:
    params = {"page": page} if page > 1 else None
    response = session.get(url, params=params, timeout=30)
    response.raise_for_status()
    return response.text


def _has_next_page(soup: BeautifulSoup) -> bool:
    """Determine if a subsequent page of reviews is available."""

    # Common patterns: rel="next" link tags, pagination components with an active "next" link
    if soup.select_one("link[rel=next], a[rel=next]"):
        return True

    next_buttons = soup.select(".pagination .next, .pager .next, a[aria-label='Next']")
    for button in next_buttons:
        if isinstance(button, Tag) and "disabled" not in button.get("class", []):
            return True

    return False


def scrape_reviews(url: str, max_pages: int = 10) -> List[Review]:
    """Fetch reviews from the provided Klantenvertellen review URL."""
    all_reviews: List[Review] = []

    with requests.Session() as session:
        session.headers.update(_default_headers())

        for page in range(1, max_pages + 1):
            try:
                html = _fetch_page(session, url, page)
            except requests.RequestException:
                break

            soup = BeautifulSoup(html, "html.parser")
            current_reviews = _parse_ld_json(soup)
            if not current_reviews:
                current_reviews = _parse_dom_reviews(soup)

            if not current_reviews:
                break

            all_reviews.extend(current_reviews)

            # Prefer explicit pagination signals; fall back to item count heuristics
            if not _has_next_page(soup) or len(current_reviews) < 10:
                break

    return all_reviews


def save_to_excel(reviews: List[Review], filename: str = "merk_echt_reviews.xlsx") -> bool:
    """Persist scraped reviews to an Excel spreadsheet."""
    if not reviews:
        return False

    df = pd.DataFrame([review.to_dict() for review in reviews])
    df.to_excel(filename, index=False, sheet_name="Reviews")
    return True


def main() -> int:
    parser = argparse.ArgumentParser(description="Scrape Klantenvertellen reviews into Excel")
    parser.add_argument(
        "--url",
        default="https://www.klantenvertellen.nl/reviews/1039690/merk_echt",
        help="Klantenvertellen review page URL",
    )
    parser.add_argument(
        "--pages",
        type=int,
        default=10,
        help="Maximum number of pages to paginate through",
    )
    parser.add_argument(
        "--output",
        default="merk_echt_reviews.xlsx",
        help="Output Excel filename",
    )
    args = parser.parse_args()

    reviews = scrape_reviews(args.url, max_pages=args.pages)

    if not reviews:
        print("No reviews were scraped. Please check the URL or selectors.")
        return 1

    save_to_excel(reviews, args.output)
    print(f"Saved {len(reviews)} reviews to {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
