"""Ajoute une photo Unsplash à la banque d'images d'une édition.

Usage :
  python ecommerce_watch/unsplash.py <dossier_assets> <nom> <url_image> "<photographe>"

- <url_image> : lien direct https://images.unsplash.com/photo-... relevé sur la page de la photo.
  Les photos Unsplash+ (plus.unsplash.com, premium_photo) sont refusées : elles ne sont pas libres.
- Écrit <dossier_assets>/images/<nom>.jpg (2000 px de large) et met à jour <dossier_assets>/index.json
  avec le crédit « Photo : <photographe> / Unsplash », repris automatiquement sous la photo.
- Affiche la taille de l'image pour vérifier qu'elle est exploitable.
"""

from __future__ import annotations

import json
import re
import sys
import urllib.request
from pathlib import Path

from PIL import Image


def main(folder: str, name: str, url: str, photographer: str) -> int:
    if "plus.unsplash.com" in url or "premium_photo" in url:
        print("Refusé : photo Unsplash+ (premium), non libre de droits.")
        return 1
    m = re.match(r"https://images\.unsplash\.com/(photo-[\w-]+)", url)
    if not m:
        print("Refusé : il faut un lien direct https://images.unsplash.com/photo-...")
        return 1
    out = Path(folder)
    (out / "images").mkdir(parents=True, exist_ok=True)
    dest = out / "images" / f"{name}.jpg"
    src = f"https://images.unsplash.com/{m.group(1)}?fm=jpg&q=82&w=2000"
    with urllib.request.urlopen(urllib.request.Request(src, headers={"User-Agent": "veille-ecommerce"}), timeout=60) as r:
        dest.write_bytes(r.read())
    with Image.open(dest) as im:
        w, h = im.size
    index_path = out / "index.json"
    index = json.loads(index_path.read_text(encoding="utf-8")) if index_path.exists() else {"images": []}
    index["images"] = [i for i in index["images"] if i["file"] != f"images/{name}.jpg"]
    index["images"].append({"file": f"images/{name}.jpg", "source": "Unsplash", "license": "Unsplash License",
                            "credit": f"Photo : {photographer} / Unsplash", "url": src})
    index_path.write_text(json.dumps(index, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"OK images/{name}.jpg {w}x{h}")
    return 0


if __name__ == "__main__":
    if len(sys.argv) != 5:
        sys.exit(__doc__)
    sys.exit(main(*sys.argv[1:]))
