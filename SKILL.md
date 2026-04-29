---
name: oh-my-pptx
description: Convert one-page, multi-page, or responsive HTML slide designs into a single native PPTX deck with DeckFlow/deckops. Use when Codex or Claude Code needs to prepare generated .html slides, split them into accurate one-page HTML inputs, inspect the split output, choose a PPTX-friendly canvas size, and submit all HTML pages to DeckFlow batch HTML-to-PPTX conversion. Supports DeckFlow's --need-embed-fonts option, defaulting to false.
---

# Oh My PPTX

## Overview

Use DeckFlow's `deckops` CLI for HTML-to-PPTX conversion. The bundled wrapper prepares HTML for DeckFlow by splitting pages, normalizing the canvas, inlining local assets, removing nonessential scripts, inspecting the generated single-page HTML files, and then submitting all pages in one batch conversion.

DeckFlow now supports multiple HTML inputs in one HTML-to-PPTX request, so the default path is batch conversion, not per-page conversion plus `join`.

## Requirements

- Node.js 18 or newer.
- Latest `deckops` with batch HTML-to-PPTX support.
- DeckFlow authentication configured with `deckops login`, or with `deckops config set-token` and `deckops config set-space`.

Install or upgrade:

```bash
npm install -g deckops@latest
deckops login
```

Confirm the CLI exposes these options:

```bash
deckops convert --help
```

Look for `convert <input-files...> --to pptx`, `--width`, `--height`, and `--need-embed-fonts`.

## Workflow

1. Create or locate an `.html` file that represents the slide deck or responsive page.
2. Run an inspection pass first:

```bash
node /path/to/oh-my-pptx/scripts/html-to-pptx.mjs input.html --inspect-only --keep-temp --json
```

3. Check the reported slide count, split mode, canvas size, warnings, and temporary single-page HTML files.
4. Convert only after the inspection is acceptable:

```bash
node /path/to/oh-my-pptx/scripts/html-to-pptx.mjs input.html --out output.pptx --timeout 600
```

The wrapper submits all generated pages in order:

```bash
deckops --json convert slide-001.html slide-002.html --to pptx --width 1920 --height 1080 --need-embed-fonts false
```

The user-facing result is only the final PPTX. Do not return intermediate HTML files or inspection JSON unless explicitly asked for debugging.

## Splitting Rules

- Prefer one page-sized slide element per slide, such as `.slide`, `section.slide`, `[data-slide]`, `[data-slide-index]`, or `[data-page]`.
- The wrapper preserves the full original DOM for selector-based splits and hides non-active slides. This keeps selectors such as `:nth-child`, `:nth-of-type`, and sibling selectors stable.
- The active slide is marked with `visible`, which makes common reveal-animation decks render their final visible state without waiting for IntersectionObserver JavaScript.
- Local CSS, images, and fonts are inlined by default so DeckFlow can render uploaded HTML without access to the user's filesystem.
- Scripts are stripped from generated pages by default to remove interactive controls, exporters, and navigation logic before upload. Use `--script-mode keep` only when the slide content is actually rendered by JavaScript.
- If `.slide` uses `100vw` and `100vh`/`100dvh`, the wrapper treats it as a responsive slide deck and uses a `1920x1080` PPTX canvas unless you override it.
- If the source is a long responsive page instead of a slide deck, use a fixed PPTX canvas and scroll splitting:

```bash
node /path/to/oh-my-pptx/scripts/html-to-pptx.mjs page.html \
  --split-mode scroll \
  --page-count 4 \
  --canvas 1920x1080 \
  --inspect-only \
  --keep-temp
```

- If dimensions are ambiguous, pass `--width <px> --height <px>` or `--canvas 16:9`.
- Use `--strict` when warnings should block upload.

## Font Embedding

DeckFlow's `--need-embed-fonts` defaults to `false`. Keep the default unless the user explicitly needs fonts embedded into the generated PPTX:

```bash
node /path/to/oh-my-pptx/scripts/html-to-pptx.mjs input.html --out output.pptx --need-embed-fonts true
```

## HTML Authoring Guidelines

- Use fixed slide thinking when possible: `.slide { width: 1920px; height: 1080px; }`.
- Prefer real HTML text for titles and body copy; avoid baking text into images.
- Use web-safe fonts or include local font files via `@font-face`; the wrapper will inline local font files when possible.
- Avoid animations, hover states, lazy loading, and JavaScript-only rendering for final conversion.
- Reveal animations are acceptable when the final state is controlled by a `visible` ancestor class.
- For responsive long pages, inspect in a browser first, choose a PPTX canvas, then use `--split-mode scroll --page-count N`.

## Troubleshooting

- `NO_SPACE_ID` or authentication errors: run `deckops login` or configure token and space.
- Missing `--need-embed-fonts` or multi-input support: upgrade with `npm install -g deckops@latest`.
- No local `.pptx` is produced: rerun with `--inspect-only --keep-temp --json` to inspect split HTML and warnings.
- Missing content that is rendered by JavaScript: rerun with `--script-mode keep`, or pre-render the content into static HTML before conversion.
- Layout drift: use selector-based slides, fix the canvas explicitly, inline assets, strip scripts where possible, and inspect the generated single-page HTML before uploading.
