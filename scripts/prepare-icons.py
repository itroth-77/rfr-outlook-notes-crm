"""Redimensionnement technique du logo fourni, sans recadrage ni generation IA.

Execution locale seulement (Pillow) ; les PNG sont versionnes, aucun nouvel
outil/dependance n'est requis par le build GitHub Pages.
"""
from pathlib import Path
from PIL import Image

ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "assets/rdv-notes-logo-source.png"
SIZES = (16, 25, 32, 48, 50, 64, 75, 80, 96, 128, 144)


def main():
    with Image.open(SOURCE) as source:
        if source.width != source.height or source.width < max(SIZES):
            raise ValueError("logo_source_must_be_square_and_large_enough")
        rgba = source.convert("RGBA")
        for size in SIZES:
            target = ROOT / "public/assets" / f"rdv-notes-logo-{size}.png"
            rgba.resize((size, size), Image.Resampling.LANCZOS).save(target, "PNG", optimize=True)
    print("Onze tailles PNG preparees, logo complet conserve.")


if __name__ == "__main__":
    main()
