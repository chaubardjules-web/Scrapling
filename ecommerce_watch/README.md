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

## Newsletter PowerPoint

`newsletter_deck.py` transforme un fichier de contenu JSON en deck PowerPoint, toujours sur le même gabarit : charte Deloitte (noir, blanc, vert), police Aptos, page de garde avec sommaire, puis une slide par news (1 ou 2 gros chiffres, infos clés, pourquoi ça compte, graphique natif ou icône). La mise en page s'appuie sur [pptx-mcp-server](https://github.com/knorq-ai/pptx-mcp-server), installé à une version figée pour que le rendu ne bouge pas.

```bash
pip install pyyaml "pptx-mcp-server @ git+https://github.com/knorq-ai/pptx-mcp-server@d5c41228fcc81e7a0ffdd7b870500a1d25136d0f"
python ecommerce_watch/newsletter_deck.py ecommerce_watch/exemple_contenu.json veille.pptx
python ecommerce_watch/newsletter_deck.py --icons   # icônes disponibles
```

Le format du JSON est décrit en tête de `newsletter_deck.py` ; `exemple_contenu.json` sert de référence.
