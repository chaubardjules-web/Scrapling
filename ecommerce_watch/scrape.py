"""Veille e-commerce : scrape les sources de sources.yaml avec Scrapling
et écrit les articles récents dans data/latest.json (+ une archive datée).

Usage : python ecommerce_watch/scrape.py
"""

from __future__ import annotations

import json
import re
import sys
import time
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta, timezone
from email.utils import parsedate_to_datetime
from pathlib import Path
from urllib.parse import urljoin

import yaml

from scrapling.fetchers import Fetcher
from scrapling.parser import Selector

ROOT = Path(__file__).resolve().parent
DATA = ROOT / "data"
ARTICLE_SELECTORS = ("article p", "main p", "[class*='content'] p", "[class*='article'] p", "p")


# ---------------------------------------------------------------- helpers

def clean_html(fragment: str) -> str:
    if not fragment:
        return ""
    if "<" not in fragment:
        return re.sub(r"\s+", " ", fragment).strip()
    text = Selector(content=f"<div>{fragment}</div>").get_all_text(separator=" ", strip=True)
    return re.sub(r"\s+", " ", str(text)).strip()


def parse_date(raw: str | None) -> datetime | None:
    if not raw:
        return None
    raw = raw.strip()
    try:
        dt = parsedate_to_datetime(raw)
    except (TypeError, ValueError):
        try:
            dt = datetime.fromisoformat(raw.replace("Z", "+00:00"))
        except ValueError:
            return None
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return dt.astimezone(timezone.utc)


def local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1].lower()


def child_text(node: ET.Element, *names: str) -> str:
    for child in node:
        if local(child.tag) in names and (child.text or "").strip():
            return child.text.strip()
    return ""


def child_link(node: ET.Element) -> str:
    for child in node:
        if local(child.tag) != "link":
            continue
        if child.get("href") and child.get("rel", "alternate") == "alternate":
            return child.get("href")
        if (child.text or "").strip():
            return child.text.strip()
    return ""


def child_source(node: ET.Element) -> str:
    for child in node:
        if local(child.tag) == "source" and (child.text or "").strip():
            return child.text.strip()
    return ""


def fetch(url: str):
    return Fetcher.get(url, timeout=30, stealthy_headers=True, follow_redirects=True, retries=2)


# ---------------------------------------------------------------- parsers

XML_ENTITIES = {"amp", "lt", "gt", "quot", "apos"}


def _xml_root(body: bytes):
    """Parse strict, puis tolérant (flux mal formés : entités HTML, BOM, octets parasites)."""
    try:
        return ET.fromstring(body)
    except ET.ParseError:
        import html
        from lxml import etree
        text = body.decode("utf-8", errors="replace").lstrip("\ufeff \r\n\t")
        text = re.sub(r"&([A-Za-z][A-Za-z0-9]*);",
                      lambda m: m.group(0) if m.group(1) in XML_ENTITIES
                      else html.escape(html.unescape(m.group(0)), quote=False),
                      text)
        text = re.sub(r"&(?!(?:[A-Za-z][A-Za-z0-9]*|#\d+|#x[0-9A-Fa-f]+);)", "&amp;", text)
        root = etree.fromstring(text.encode("utf-8"),
                                etree.XMLParser(recover=True, huge_tree=True, resolve_entities=False))
        if root is None:
            raise
        etree.strip_tags(root, etree.Comment, etree.ProcessingInstruction)
        return ET.fromstring(etree.tostring(root))


def parse_rss(body: bytes) -> list[dict]:
    root = _xml_root(body)
    items = []
    for node in root.iter():
        if local(node.tag) not in ("item", "entry"):
            continue
        items.append({
            "title": clean_html(child_text(node, "title")),
            "url": child_link(node),
            "published": parse_date(child_text(node, "pubdate", "published", "updated", "date")),
            "summary": clean_html(child_text(node, "description", "summary", "content", "encoded"))[:600],
            "publisher": child_source(node),
        })
    return items


def parse_html(page, src: dict) -> list[dict]:
    items = []
    for block in page.css(src["item"]):
        title = block.css(src["title"]).get()
        link = block.css(src["link"]).get()
        if not title or not link:
            continue
        date_sel = src.get("date")
        items.append({
            "title": clean_html(str(title)),
            "url": urljoin(src["url"], str(link)),
            "published": parse_date(str(block.css(date_sel).get())) if date_sel else None,
            "summary": "",
            "publisher": "",
        })
    return items


def article_excerpt(url: str, max_chars: int) -> str:
    page = fetch(url)
    if page.status != 200:
        return ""
    for sel in ARTICLE_SELECTORS:
        paras = [str(p.get_all_text(separator=" ", strip=True)) for p in page.css(sel)]
        paras = [p for p in paras if len(p) > 60]
        if len(paras) >= 2:
            return re.sub(r"\s+", " ", " ".join(paras))[:max_chars]
    return ""


def matches(item: dict, keywords: list[str]) -> bool:
    haystack = f"{item['title']} {item['summary']}".lower()
    return any(re.search(rf"\b{re.escape(k.lower())}", haystack) for k in keywords)


# ---------------------------------------------------------------- main

def main() -> int:
    config = yaml.safe_load((ROOT / "sources.yaml").read_text(encoding="utf-8"))
    settings = config.get("settings", {})
    window = timedelta(days=settings.get("window_days", 5))
    per_source = settings.get("max_items_per_source", 25)
    max_chars = settings.get("article_chars", 1800)
    keywords = config.get("keywords", [])
    now = datetime.now(timezone.utc)
    cutoff = now - window

    articles, report, seen = [], [], set()
    for src in config["sources"]:
        status = {"source": src["name"], "url": src["url"], "ok": False, "kept": 0}
        try:
            page = fetch(src["url"])
            if page.status != 200:
                raise RuntimeError(f"HTTP {page.status}")
            items = parse_rss(page.body) if src.get("type", "rss") == "rss" else parse_html(page, src)
            kept = 0
            for item in items:
                if kept >= per_source:
                    break
                if not item["url"] or item["url"] in seen:
                    continue
                if item["published"] and item["published"] < cutoff:
                    continue
                if src.get("filter") and not matches(item, keywords):
                    continue
                seen.add(item["url"])
                # Google News suffixe les titres par " - Éditeur"
                if item["publisher"] and item["title"].endswith(f" - {item['publisher']}"):
                    item["title"] = item["title"][: -len(item["publisher"]) - 3].strip()
                excerpt = ""
                if src.get("fetch_article", True):
                    try:
                        excerpt = article_excerpt(item["url"], max_chars)
                    except Exception:
                        pass
                    time.sleep(0.5)
                articles.append({
                    "title": item["title"],
                    "url": item["url"],
                    "source": src["name"],
                    "publisher": item["publisher"] or src["name"],
                    "theme": src.get("theme", ""),
                    "published": item["published"].isoformat() if item["published"] else None,
                    "summary": item["summary"],
                    "excerpt": excerpt,
                })
                kept += 1
            status.update(ok=True, kept=kept)
        except Exception as exc:
            status["error"] = f"{type(exc).__name__}: {exc}"[:200]
        report.append(status)
        print(f"[{'OK' if status['ok'] else 'KO'}] {src['name']}: {status.get('kept', 0)} articles"
              + (f" ({status['error']})" if "error" in status else ""))

    articles.sort(key=lambda a: a["published"] or "", reverse=True)
    output = {
        "generated_at": now.isoformat(),
        "window_start": cutoff.isoformat(),
        "count": len(articles),
        "sources": report,
        "articles": articles,
    }
    DATA.mkdir(exist_ok=True)
    payload = json.dumps(output, ensure_ascii=False, indent=2)
    (DATA / "latest.json").write_text(payload, encoding="utf-8")
    (DATA / f"{now:%Y-%m-%d}.json").write_text(payload, encoding="utf-8")
    archives = sorted(p for p in DATA.glob("????-??-??.json"))
    for old in archives[:-24]:  # ~3 mois d'historique à 2 runs/semaine
        old.unlink()
    print(f"\n{len(articles)} articles écrits dans {DATA / 'latest.json'}")

    ok_sources = sum(1 for r in report if r["ok"])
    return 0 if ok_sources else 1


if __name__ == "__main__":
    sys.exit(main())
