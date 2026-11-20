# PPTX Editor for VSCode

A VSCode extension built on [`@fefeding/ppt-parser`](https://github.com/fefeding/pptx-parser) that lets you **open and edit `.pptx` files visually** — directly inside the editor.

> **Monorepo note**: This extension now lives in the `vscode/` subdirectory of the
> [pptx-parser repository](https://github.com/fefeding/pptx-parser). It no longer installs
> `@fefeding/ppt-parser` from npm — the build compiles the parser's TypeScript source from
> `../src` and inlines it into the extension bundle, so the published VSIX is self-contained.
> Releases are triggered by `vsix-v*` tags in the pptx-parser repo.

![PPTX editor screenshot](media/screenshot.png)

## Features

- **Default `.pptx` handler**: Double-click any `.pptx` in the Explorer to open it with the built-in visual editor (no extra software required).
- **PowerPoint-style ribbon**: a single compact header row with the file name, quick access
  (undo / redo / slide show), the preview ⇄ edit switch and a collapse toggle. Edit mode
  adds the Home / Insert / Design / Transitions / Animations / Slide Show / View / Help tabs whose
  grouped controls are hidden by default and appear when a tab is clicked. Preview mode replaces
  the tab ribbon with a compact ☰ shortcut menu (slide show, fit / actual size, shortcuts); zoom
  lives only in the status bar. Saving is handled by VS Code itself (`Ctrl/Cmd+S`), and the header
  dot reflects the document's real unsaved state.
- **Localized UI**: follows VS Code's display language (English / 简体中文 bundled).
- **Slide preview**: Renders every slide at its original aspect ratio via `pptxToHtml` (most faithful). Left sidebar thumbnail navigation with zoom-to-fit and grid view.
- **Visual editing**:
  - Click to select elements on the canvas; drag to move; drag handles to resize; drag the rotation grip to rotate.
  - Properties panel: position / size / rotation, text (content / font size / color / alignment / bold / italic / underline), shapes (fill / border), image replacement.
  - **Preset shapes render exactly like PowerPoint**: the canvas uses the same geometry formulas as the
    preview (moon, sun, cloud, donut, gear, heart, lightning, cube, folded corner, callouts …), including
    solid **and gradient** fills — no more CSS approximations that turn a moon into a circle.
  - **Paragraph refinement**: left / right indent, hanging indent, space before / after, list level
    (promote / demote), text-box padding and auto-fit — all written back as the corresponding
    OOXML paragraph properties (`a:pPr@marL/marR/indent/lvl`, `a:spcBef/a:spcAft`, `a:bodyPr` insets
    and `a:normAutofit`).
  - **Connectors & snapping**: the shape picker offers straight / bent / curved connectors. Drag an
    endpoint near a shape and it snaps to that shape's connection point; the connection is stored as
    the native `a:stCxn` / `a:endCxn` relation, so the line keeps following the shape in PowerPoint.
  - **Editable SmartArt**: double-click a SmartArt graphic to edit the text of every node. The text
    is written back into both the SmartArt data part and its cached drawing part, so the original
    layout, colours and styling are kept.
  - The element list is a collapsible drawer in the panel header (remembered across sessions): collapsed by default so the properties stay in view, with a name filter and its own scroll box for slides that carry dozens of elements.
  - Slide operations: add blank slide, duplicate, delete, drag thumbnails to reorder, set background color, lock, hide, speaker notes.
  - **Two slide-panel views**: rendered thumbnails, or a text **outline** — one row per slide showing its title, with hidden slides struck through and an `N` marker for slides that carry speaker notes. Right-click a row for the same slide menu as a thumbnail; the choice is remembered.
  - **Virtualized slide panel**: only the thumbnails near the viewport are drawn, and a thumbnail is
    redrawn only when its slide actually changes — editing a slide in a long deck no longer
    re-renders every thumbnail (measured ~10× faster navigation on a chart-heavy 21-slide deck).
  - Insert elements: text box, rectangle, image (from local file).
  - Alignment / distribution / grouping / z-order / find-replace via the editor core actions.
- **Presentation mode**: `F5` (from the beginning) or `Shift+F5` (from the current slide) for fullscreen slideshow. Navigate with arrow keys, space or click; `Esc` exits.
- **Save & export**: saving is delegated to VS Code (`Ctrl/Cmd+S`) — the header dot tracks the real
  unsaved state; "Export Copy…" saves a copy; "Discard Changes & Reload" reverts to disk.
- **Export images / PDF**: the View tab (edit mode) and the preview ☰ menu can rasterize the
  current slide or the whole deck to PNG at 2× slide size — ECharts charts, SmartArt shapes and
  video poster frames included — and save them individually or assemble them into a PDF. Slides
  are rendered offscreen from the editor model and rasterized with `html-to-image`.
- **Undo / Redo**: `Ctrl/Cmd+Z`, `Ctrl/Cmd+Y`. Press `?` for the shortcut reference.

## Development

```bash
# Run from the <pptx-parser-repo>/vscode directory. The parser source (../src) must be present.
npm install            # installs build/runtime deps; parser itself is compiled from source, not npm
npm run compile        # esbuild bundles dist/extension.js and dist/webview.js (+ inlined parser)
npm run watch          # watch mode (alias: npm run dev)
npm run verify:convert # end-to-end conversion check (no VS Code instance needed)
npm run package        # build the production bundle and package a .vsix (vsce package)
npm run publish        # publish to the VS Code Marketplace (vsce publish)
```

Press `F5` in VSCode to launch an Extension Development Host, then open any `.pptx` file.
The default zoom for both preview and edit canvases is "fit to window"; use the status-bar
zoom controls (or `Ctrl/Cmd` + wheel on the canvas) for manual zoom.

## Localization

The editor UI ships in `src/webview/locales/*.ts` and follows VS Code's display language
(`vscode.env.language`, delivered with the webview HTML and the `init` message). `en.ts` is the
canonical key set — other languages are typed as `Record<keyof typeof en, string>`, so a missing
translation fails `npx tsc --noEmit`. To add a language, copy `en.ts`, translate it, and register it
in `BUNDLES` (`src/webview/i18n.ts`); `zh-*` locales fall back to `zh-cn` until registered.

Host-side strings (command titles, save dialogs, notifications) use VS Code's `vscode.l10n`:
`package.nls.json` holds the defaults for `package.json` contributions, `l10n/bundle.l10n.json`
holds the runtime defaults, and `l10n/<lang>.json` / `l10n/<lang>.bundle.l10n.json` the translations.

## Notes & Limitations

- Editing renders straight from the editor document model provided by `@fefeding/ppt-parser`
  (its headless editor core: `createStore` / `docFromPptx` / `createActions` / `renderSlideInto`).
  On save the document is serialized back to PPTX. **Content elements** — text boxes, shapes,
  images, charts, tables — are editable and preserved, and because editing renders from the model
  (not a byte round-trip) it never loses fidelity such as picture or pattern fills.
- **Slide masters, layouts and themes survive the round-trip.** Their original XML parts are carried
  through untouched (`__rawXml`), including multi-theme decks: every `themeN.xml` /
  `slideMasterN.xml` / `slideLayoutN.xml` is written back byte-for-byte and each slide keeps its
  original layout binding (verified on samples with 2 themes / 2 masters / 24 layouts).
- The preview mode uses `pptxToHtml` for the most faithful rendering of the original file.
- Charts are rendered with ECharts (the same renderer the parser's preview uses), including 3D
  charts where supported by the source.
- Complex objects such as OLE and equations are retained as raw OOXML (`__raw`) but are not visually
  editable. SmartArt is retained the same way, with the addition that the **text of its nodes can be
  edited** (double-click the graphic) — the change is written into the diagram parts in place, so the
  original layout and colours stay untouched.
