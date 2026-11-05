# Veille e-commerce

Fork de [Scrapling](https://github.com/D4Vinci/Scrapling) qui alimente une newsletter bi-hebdomadaire sur les tendances e-commerce (agentic commerce, GEO, social commerce, nouveaux outils, retail media, beauty/luxe).

## Fonctionnement

1. Le lundi et le jeudi à l'aube, le workflow `.github/workflows/ecommerce-watch.yml` installe Scrapling depuis ce repo et lance `scrape.py`.
2. Le script récupère les sources listées dans `sources.yaml` (flux RSS, recherches Google News, pages HTML), garde les articles des 5 derniers jours, filtre par mots-clés les médias généralistes et récupère un extrait du texte de chaque article.
3. Le résultat est commité dans `data/latest.json`, avec une archive datée (environ 3 mois d'historique).
4. À 7h40, la tâche planifiée Claude lit `latest.json`, complète si besoin par une recherche web et rédige la newsletter.

## Personnaliser

- Ajouter ou retirer une source : éditer `sources.yaml`. Une source qui échoue est simplement ignorée, son statut apparaît dans `latest.json > sources`.
- Changer la fenêtre ou la taille des extraits : bloc `settings` de `sources.yaml`.
- Lancer un run à la main : onglet Actions > Veille e-commerce > Run workflow.

## En local

```bash
pip install -e ".[fetchers]" pyyaml
python ecommerce_watch/scrape.py
```
