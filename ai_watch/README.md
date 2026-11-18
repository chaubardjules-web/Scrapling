# Veille IA — Taskforce IA Deloitte

Veille hebdomadaire sur l'IA, publiée le mercredi pour la taskforce IA : modèles et labs, agents et outils, IA en entreprise, régulation et souveraineté, acteurs émergents (startups, levées, YC, IPO). Même moteur et même chaîne que la veille e-commerce (`ecommerce_watch/`).

## Fonctionnement

1. Le mercredi à l'aube, `.github/workflows/ai-watch.yml` lance `ecommerce_watch/scrape.py --config ai_watch/sources.yaml --data ai_watch/data`.
2. Le script garde les articles des 8 derniers jours, filtre par mots-clés (mots entiers) les médias généralistes et récupère un extrait de chaque article.
3. Le résultat est commité dans `data/latest.json`, avec une archive datée (13 semaines).
4. Vers 8h10, la tâche planifiée Claude lit `latest.json`, complète par la recherche web et Undermind, produit le deck, le podcast « AI FM » et alimente le hub Notion « Taskforce IA — Hub Deloitte ».

## Audit des sources

Le 1er de chaque mois, `.github/workflows/ai-source-audit.yml` mesure rythme et pertinence de chaque source de `audit/candidates.yaml` et écrit `audit/latest.json` et `audit/latest.md`. On ajuste `sources.yaml` à partir de ce classement.

## En local

```bash
pip install -e ".[fetchers]" pyyaml
python ecommerce_watch/scrape.py --config ai_watch/sources.yaml --data ai_watch/data
```
