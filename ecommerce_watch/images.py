"""Construit la banque d'images libres de droits de la newsletter via l'API Openverse.

Lit data/latest.json (articles du jour) et image_topics.yaml, cherche des photos
sous licence libre à usage commercial, les redimensionne et écrit :
  <out>/images/<slug>-<n>.jpg
  <out>/index.json   (requête, fichier, titre, auteur, licence, crédit, page source)

Usage : python ecommerce_watch/images.py <dossier_sortie>
"""

from __future__ import annotations

import io
import json
import re
import sys
import time
from pathlib import Path
from urllib.parse import urlencode

import requests
import yaml
from PIL import Image

ROOT = Path(__file__).resolve().parent
API = "https://api.openverse.org/v1/images/"
HEADERS = {"User-Agent": "ecommerce-watch-newsletter/1.0 (github.com/chaubardjules-web/Scrapling)"}
LICENSE_LABEL = {"cc0": "CC0", "pdm": "Domaine public", "by": "CC BY", "by-sa": "CC BY-SA"}


def slugify(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", "-", value.lower()).strip("-")[:40]


def search(query: str, per_query: int, min_width: int) -> list[dict]:
    params = {"q": query, "license": "cc0,pdm,by,by-sa", "license_type": "commercial",
              "page_size": 20, "mature": "false", "aspect_ratio": "wide"}
    for attempt in range(3):
        r = requests.get(f"{API}?{urlencode(params)}", headers=HEADERS, timeout=30)
        if r.status_code == 429:
            time.sleep(20 * (attempt + 1))
            continue
        r.raise_for_status()
        results = r.json().get("results", [])
        good = [x for x in results if (x.get("width") or 0) >= min_width and x.get("url")]
        return good[: per_query * 3]  # marge pour les téléchargements qui échouent
    return []


def credit(item: dict) -> str:
    lic = LICENSE_LABEL.get(item.get("license", ""), item.get("license", "").upper())
    if item.get("license") in ("by", "by-sa") and item.get("license_version"):
        lic = f"{lic} {item['license_version']}"
    creator = (item.get("creator") or "").strip()
    source = (item.get("source") or item.get("provider") or "").replace("_", " ").title()
    parts = [p for p in (creator, source) if p]
    return f"Photo : {' / '.join(parts)}, {lic}" if parts else f"Photo : {lic}"


def download(item: dict, dest: Path, max_width: int) -> bool:
    try:
        r = requests.get(item["url"], headers=HEADERS, timeout=40)
        r.raise_for_status()
        img = Image.open(io.BytesIO(r.content)).convert("RGB")
        if img.width > max_width:
            img = img.resize((max_width, round(img.height * max_width / img.width)), Image.LANCZOS)
        img.save(dest, "JPEG", quality=80, optimize=True, progressive=True)
        return True
    except Exception as exc:
        print(f"    échec téléchargement : {exc}"[:160])
        return False


def main(out_dir: str) -> int:
    cfg = yaml.safe_load((ROOT / "image_topics.yaml").read_text(encoding="utf-8"))
    s = cfg.get("settings", {})
    per_query, min_w, max_w = s.get("per_query", 3), s.get("min_width", 1000), s.get("max_width", 1600)

    latest = ROOT / "data" / "latest.json"
    corpus = ""
    if latest.exists():
        data = json.loads(latest.read_text(encoding="utf-8"))
        corpus = " ".join(f"{a['title']} {a.get('summary', '')}" for a in data.get("articles", [])).lower()

    queries = []
    for b in cfg.get("brands", []):
        if any(m.lower() in corpus for m in b["match"]):
            queries.append({"id": slugify(b["match"][0]), "query": b["query"], "kind": "marque"})
    for t in cfg.get("topics", []):
        queries.append({"id": t["id"], "query": t["query"], "kind": "sujet"})

    out = Path(out_dir)
    (out / "images").mkdir(parents=True, exist_ok=True)
    index, seen = [], set()
    for q in queries:
        print(f"[{q['kind']}] {q['query']}")
        try:
            results = search(q["query"], per_query, min_w)
        except Exception as exc:
            print(f"    recherche en échec : {exc}"[:160])
            continue
        kept = 0
        for item in results:
            if kept >= per_query:
                break
            if item["id"] in seen:
                continue
            dest = out / "images" / f"{q['id']}-{kept + 1}.jpg"
            if not download(item, dest, max_w):
                continue
            seen.add(item["id"])
            kept += 1
            index.append({
                "file": f"images/{dest.name}", "query": q["query"], "topic": q["id"], "kind": q["kind"],
                "title": item.get("title", ""), "creator": item.get("creator", ""),
                "license": item.get("license", ""), "license_version": item.get("license_version", ""),
                "credit": credit(item), "credit_required": item.get("license") in ("by", "by-sa"),
                "landing_page": item.get("foreign_landing_url", ""),
            })
        print(f"    {kept} image(s)")
        time.sleep(3)

    (out / "index.json").write_text(json.dumps({"generated_at": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
                                                "images": index}, ensure_ascii=False, indent=2),
                                    encoding="utf-8")
    print(f"\n{len(index)} images dans {out}")
    return 0 if index else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv[1] if len(sys.argv) > 1 else "newsletter-assets"))
