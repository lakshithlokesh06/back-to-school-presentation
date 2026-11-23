# Back to School presentation scaffold

Editable three-slide PowerPoint scaffold for the Communication Skills assignment. This phase establishes the slide structure, shared visual theme, copy source, and placeholder elements; it is not the final polished design.

## Generate

```sh
npm install
npm run build
```

The generated deck is `back-to-school-scaffold.pptx` in the repository root. The build uses the locally bundled `pptxgenjs` package pinned in `package.json`.

## Edit

- Change slide wording and placeholder descriptions in `presentation/content.js`.
- Adjust shared colors, type, page sizing, and helper layouts in `presentation/theme.js`.
- Tune slide-level positions and components in `presentation/generate.js`.

The deck contains exactly three slides: Opening, Body (the three selected points), and Close with the requested call to action. Speaker notes/scripts are intentionally not included.
