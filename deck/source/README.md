# Deck source

The canonical editable deliverable is `../current/Nana-Pitch-Deck-v3.pptx`.

`build.mjs` generated v2. `edit-v3.mjs` imports that deck and applies the landing-aligned art direction: atmospheric textures, glass panels, quiet headers, semantic accent colors and the bubble-based expansion diagram.

`content-plan.txt` records the narrative and editorial rules.

`build_pdf_v3.py` creates the visually faithful v3 PDF from high-resolution rendered slide PNGs. This avoids font substitution on machines without Fredoka and Nunito. The original `build_pdf.py` remains as the v2 source.

If editing directly in PowerPoint:

1. Keep the 16:9 canvas.
2. Install Fredoka and Nunito.
3. Preserve source notes for external facts.
4. Export a PDF and compare every page to the PPTX before sharing.
