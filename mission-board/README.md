# Nana Mission Board

Final files:

- `Nana-Mission-Board.pptx`: editable source.
- `Nana-Mission-Board.pdf`: visually faithful shareable export.

## Regenerate the PPTX

Run this exact command from the repository root:

```sh
node mission-board/source/build.mjs
```

The command writes `mission-board/Nana-Mission-Board.pptx`. The included PDF is
the shareable export validated against all 14 slides.

The generator depends on OpenAI's internal `@oai/artifact-tool` runtime and is
included for layout traceability. Normal PowerPoint editing does not require it.

PowerPoint does not provide a CSS style fallback chain. Fredoka and Nunito must
be installed locally for the PPTX to render with Nana's brand typography. Get
both fonts from Google Fonts:

- https://fonts.google.com/specimen/Fredoka
- https://fonts.google.com/specimen/Nunito
