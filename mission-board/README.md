# Nana Mission Board

## Regenerate the PPTX

Run this exact command from the repository root:

```sh
node docs/mission-board/src/build.mjs
```

The command writes `docs/mission-board/Nana-Mission-Board.pptx`, imports the
PPTX again to create `docs/mission-board/preview/slides/`, and writes the
artifact-tool source renders under `docs/mission-board/preview/artifact-tool/`.

The generator depends on `@oai/artifact-tool`. A working copy is vendored under
`docs/mission-board/src/node_modules/` because this private package is not
published on npm and `npm view @oai/artifact-tool` returns 404. No package
installation step is required.

PowerPoint does not provide a CSS style fallback chain. Fredoka and Nunito must
be installed locally for the PPTX to render with Nana's brand typography. Get
both fonts from Google Fonts:

- https://fonts.google.com/specimen/Fredoka
- https://fonts.google.com/specimen/Nunito
