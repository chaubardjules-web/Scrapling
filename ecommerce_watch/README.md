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

`newsletter_deck.py` transforme un fichier de contenu JSON en deck PowerPoint, toujours sur la même charte, définie dans `charte/charte.yaml` : coins arrondis, ombres douces, vert Deloitte en accent, famille Aptos (Display pour les titres et les chiffres, Serif italique pour « Pourquoi ça compte », Narrow pour les surtitres). Structure : une page de garde avec sommaire, puis une slide par news avec 1 ou 2 gros chiffres, 3 infos clés, une phrase « Pourquoi ça compte » et un visuel (photo libre de droits ou graphique natif). La mise en page s'appuie sur [pptx-mcp-server](https://github.com/knorq-ai/pptx-mcp-server), installé à une version figée (`PPTX_MCP_SERVER_COMMIT`).

Les photos viennent de la branche `newsletter-assets`, reconstruite à chaque run du workflow par `images.py` (Openverse, plan B Wikimedia Commons ; licences CC0, domaine public, CC BY, CC BY-SA). Les crédits sont repris automatiquement sous chaque photo.

```bash
pip install pyyaml "pptx-mcp-server @ git+https://github.com/knorq-ai/pptx-mcp-server@$(cat ecommerce_watch/PPTX_MCP_SERVER_COMMIT)"
git clone --depth 1 -b newsletter-assets https://github.com/chaubardjules-web/Scrapling.git assets
python ecommerce_watch/newsletter_deck.py ecommerce_watch/exemple_contenu.json veille.pptx --assets assets
```

Le format du JSON est décrit en tête de `newsletter_deck.py` ; `exemple_contenu.json` sert de référence de ton et de longueur. Pour faire évoluer la charte, on modifie `charte/charte.yaml` ou le générateur, jamais le contenu d'une édition.
