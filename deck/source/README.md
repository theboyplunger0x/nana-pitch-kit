# Deck source

The canonical editable deliverable is `../current/Nana-Pitch-Deck-v2.pptx`.

`build.mjs` is the source used to generate the current deck with OpenAI's internal `@oai/artifact-tool`. It is included for layout traceability and structured iteration. It is not required for normal PowerPoint editing.

`content-plan.txt` records the narrative and editorial rules.

`build_pdf.py` creates the visually faithful PDF from the rendered slide PNGs. This avoids font substitution on machines without Fredoka and Nunito.

If editing directly in PowerPoint:

1. Keep the 16:9 canvas.
2. Install Fredoka and Nunito.
3. Preserve source notes for external facts.
4. Export a PDF and compare every page to the PPTX before sharing.
