---
name: oh-my-pptx
description: Convert one-page or multi-page HTML slide designs into a single native PPTX deck with the DeckFlow/deckops CLI. Use when Codex or Claude Code needs to turn generated .html slide designs, HTML/CSS presentation pages, or web-rendered slide mockups into .pptx files through DeckFlow/deckops, including splitting multi-slide HTML, detecting slide dimensions, converting each page, and joining the final deck.
---

# Oh My PPTX

## Overview

Use DeckFlow's `deckops` CLI for HTML-to-PPTX conversion. Prefer the bundled wrapper script so the agent does not have to split slides, pass dimensions, join outputs, or parse task output by hand.

## Requirements

- Node.js 18 or newer.
- A join-enabled `deckops` installed globally, or let the wrapper fall back to `npm exec -y deckops@latest`.
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
2. Use one page-sized `.slide` element per slide. The wrapper splits multi-slide HTML into single-slide HTML files because DeckFlow's HTML-to-PPTX conversion handles one slide per request.
3. Use explicit slide dimensions in the HTML. A reliable default is `1920px x 1080px`; the wrapper detects `.slide` dimensions and passes width/height to DeckFlow.
4. Keep assets local or absolute-URL addressable so DeckFlow can render them after upload.
5. Convert with the wrapper:

```bash
node /path/to/oh-my-pptx/scripts/html-to-pptx.mjs input.html --out output.pptx --timeout 300
```

For each slide, the wrapper runs a DeckFlow conversion task like:

```bash
deckops --json run convertor.html2pptx slide-001.html --param width=1920 --param height=1080
```

If the input has more than one slide, the wrapper then runs:

```bash
deckops --json join slide-001.pptx slide-002.pptx
```

The user-facing result is only the final PPTX. Do not return intermediate PPTX files or task JSON unless explicitly asked for debugging.

## HTML Authoring Guidelines

- Use inline CSS or a single local stylesheet; avoid build steps unless the user asks for them.
- Use `@page`-like fixed canvas thinking: `.slide { width: 1920px; height: 1080px; }`.
- If dimensions are ambiguous or generated dynamically, pass `--width <px> --height <px>` explicitly.
- Prefer real text elements for titles and body copy; avoid baking text into images.
- Use web-safe fonts or explicitly available fonts. If exact font fidelity matters, include font files and `@font-face`.
- Avoid animations, hover states, and viewport-dependent layout for final conversion.
- For multiple slides, use repeated page-sized `.slide` sections and avoid content spilling outside each slide.

## Troubleshooting

- `NO_SPACE_ID` or authentication errors: run `deckops login` or configure token and space.
- Network or npm registry errors: install a join-enabled `deckops` globally before using the wrapper.
- If multi-slide conversion fails at the join step: confirm `deckops join file1.pptx file2.pptx` works in the local CLI.
- No local `.pptx` is produced: rerun with `--keep-temp --json` to inspect the temporary slide HTML/PPTX files and task summary.
- Layout drift: simplify CSS, fix slide dimensions, inline critical styles, and rerun conversion.

## Install From A Cloned Repo

From the cloned `oh-my-pptx` folder:

```bash
./scripts/install.sh --all
```
