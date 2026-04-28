---
name: oh-my-pptx
description: Convert HTML slide pages into native PPTX decks with the deckops CLI. Use when Codex or Claude Code needs to turn generated .html slide designs, HTML/CSS presentation pages, or web-rendered slide mockups into .pptx files through Deckflow/deckops, including one-off conversions and repeatable slide-generation workflows.
---

# Oh My PPTX

## Overview

Use `deckops` for HTML-to-PPTX conversion. Prefer the bundled wrapper script so the agent does not have to remember the exact CLI flags or parse task output by hand.

## Requirements

- Node.js 18 or newer.
- `deckops` installed globally, or let the wrapper fall back to `npm exec -y deckops@0.2.1`.
- Deckops authentication configured with `deckops login`, or with `deckops config set-token` and `deckops config set-space`.

Check setup:

```bash
node -v
deckops config show
```

If `deckops` is missing:

```bash
npm install -g deckops
deckops login
```

## Workflow

1. Create or locate an `.html` file that represents the slide deck.
2. Use explicit slide dimensions in the HTML. A reliable default is `1920px x 1080px` with a single `.slide` root for each page-sized slide.
3. Keep assets local or absolute-URL addressable so Deckflow can render them after upload.
4. Convert with the wrapper:

```bash
node /path/to/oh-my-pptx/scripts/html-to-pptx.mjs input.html --out output.pptx --timeout 300
```

The wrapper runs:

```bash
deckops --json convert input.html --to pptx --timeout 300
```

It saves the returned task JSON next to the requested output and downloads the first PPTX URL it can find in `task.result`. If the backend returns a completed task without a direct URL, inspect the saved `.task.json` and fetch the result URL manually.

## HTML Authoring Guidelines

- Use inline CSS or a single local stylesheet; avoid build steps unless the user asks for them.
- Use `@page`-like fixed canvas thinking: `.slide { width: 1920px; height: 1080px; }`.
- Prefer real text elements for titles and body copy; avoid baking text into images.
- Use web-safe fonts or explicitly available fonts. If exact font fidelity matters, include font files and `@font-face`.
- Avoid animations, hover states, and viewport-dependent layout for final conversion.
- For multiple slides, use repeated page-sized `.slide` sections and avoid content spilling outside each slide.

## Troubleshooting

- `NO_SPACE_ID` or authentication errors: run `deckops login` or configure token and space.
- Network or npm registry errors: install `deckops` globally before using the wrapper.
- No local `.pptx` is produced: open the `.task.json` file saved by the wrapper and look for a downloadable result URL.
- Layout drift: simplify CSS, fix slide dimensions, inline critical styles, and rerun conversion.

## Install From A Cloned Repo

From the cloned `oh-my-pptx` folder:

```bash
./scripts/install.sh --all
```
