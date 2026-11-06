"""Audit des sources de veille : rythme de publication et pertinence par rapport au périmètre.

Pour chaque source de candidates.yaml (et celles déjà actives dans sources.yaml), le script
récupère le flux avec Scrapling, puis mesure sur les articles récupérés :
  - per_day   : articles publiés par jour (sur la période couverte par le flux)
  - relevance : part des articles qui touchent le périmètre (score de mots-clés >= 2)
  - core      : part des articles qui touchent le cœur de l'offre (un mot-clé de poids 3)
  - score     : articles pertinents par semaine, avec un bonus pour le cœur de l'offre

Écrit audit/latest.json et audit/latest.md (classement), conservés dans le repo.
Usage : python ecommerce_watch/audit/source_audit.py
"""

from __future__ import annotations

import json
import re
import sys
from datetime import datetime, timezone
from pathlib import Path

import yaml

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent))
from scrape import fetch, parse_rss  # noqa: E402  (même moteur que la veille)


def weigh(text: str, keywords: dict[int, list[str]]) -> tuple[int, bool]:
    text = f" {text.lower()} "
    score, core = 0, False
    for weight, words in keywords.items():
        for w in words:
            if re.search(rf"(?<![\w]){re.escape(w.strip().lower())}(?![\w])", text):
                score += int(weight)
                core = core or int(weight) == 3
                break  # un seul hit par niveau pour éviter les sur-comptes
    return score, core


def audit(src: dict, keywords: dict) -> dict:
    row = {"name": src["name"], "url": src["url"], "family": src.get("family", ""),
           "lang": src.get("lang", ""), "ok": False}
    try:
        page = fetch(src["url"])
        if page.status != 200:
            raise RuntimeError(f"HTTP {page.status}")
        items = [i for i in parse_rss(page.body) if i.get("title")]
        if not items:
            raise RuntimeError("flux vide ou illisible")
        dates = sorted(i["published"] for i in items if i["published"])
        now = datetime.now(timezone.utc)
        span = max((now - dates[0]).total_seconds() / 86400, 1) if dates else 7
        scored = [weigh(f"{i['title']} {i['summary']}", keywords) for i in items]
        relevant = sum(1 for s, _ in scored if s >= 2)
        broad = sum(1 for s, _ in scored if s >= 1)
        core = sum(1 for _, c in scored if c)
        per_day = len(items) / span
        freshness = (now - dates[-1]).total_seconds() / 86400 if dates else None
        row.update(
            ok=True, items=len(items), span_days=round(span, 1), per_day=round(per_day, 2),
            relevance=round(relevant / len(items), 2), core=round(core / len(items), 2),
            broad=round(broad / len(items), 2),
            last_post_days=round(freshness, 1) if freshness is not None else None,
            score=round(per_day * 7 * (relevant / len(items)) + per_day * 7 * (core / len(items)) * 2, 1),
            examples=[i["title"][:110] for i, (s, c) in zip(items, scored) if c][:3],
        )
    except Exception as exc:
        row["error"] = f"{type(exc).__name__}: {exc}"[:160]
    return row


def main() -> int:
    cfg = yaml.safe_load((HERE / "candidates.yaml").read_text(encoding="utf-8"))
    keywords = {int(k): v for k, v in cfg["keywords"].items()}
    seen, rows = set(), []
    for src in cfg["candidates"]:
        if src["url"] in seen:
            continue
        seen.add(src["url"])
        row = audit(src, keywords)
        rows.append(row)
        print(f"[{'OK' if row['ok'] else 'KO'}] {row['name']}: "
              + (f"{row['per_day']}/j, pertinence {row['relevance']}, cœur {row['core']}, score {row['score']}"
                 if row["ok"] else row.get("error", "")))
    rows.sort(key=lambda r: (r["ok"], r.get("score", 0)), reverse=True)
    now = datetime.now(timezone.utc)
    (HERE / "latest.json").write_text(json.dumps({"generated_at": now.isoformat(), "sources": rows},
                                                 ensure_ascii=False, indent=2), encoding="utf-8")
    lines = [f"# Audit des sources — {now:%Y-%m-%d}", "",
             "| Source | Famille | Articles/jour | Pertinence | Cœur de l'offre | Périmètre élargi | Dernier article (j) | Score |",
             "|---|---|---|---|---|---|---|---|"]
    for r in rows:
        if r["ok"]:
            lines.append(f"| {r['name']} | {r['family']} | {r['per_day']} | {int(r['relevance']*100)} % | "
                         f"{int(r['core']*100)} % | {int(r['broad']*100)} % | {r['last_post_days']} | {r['score']} |")
        else:
            lines.append(f"| {r['name']} | {r['family']} | — | — | — | — | — | KO ({r.get('error', '')[:40]}) |")
    (HERE / "latest.md").write_text("\n".join(lines) + "\n", encoding="utf-8")
    return 0


if __name__ == "__main__":
    sys.exit(main())
