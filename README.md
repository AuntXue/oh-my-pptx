# Oh My PPTX

Oh My PPTX is an agent skill for turning HTML slide decks into editable PowerPoint files with [DeckFlow](https://deckflow.com) and the [`deckops`](https://www.npmjs.com/package/deckops) CLI.

It is made for Claude Code, Codex, and other coding agents that create slides as HTML/CSS but need to deliver a native `.pptx`.

## What It Does

- Splits one HTML deck into ordered single-slide HTML pages.
- Preserves responsive slide layouts and common reveal-animation final states.
- Checks the split pages before upload.
- Sends the pages to DeckFlow in one batch HTML-to-PPTX conversion.
- Downloads one final editable PPTX.

## Install

Install or upgrade DeckFlow CLI:

```bash
npm install -g deckops@latest
deckops login
```

Install the skill:

```bash
git clone https://github.com/AuntXue/oh-my-pptx.git
cd oh-my-pptx
./scripts/install.sh --all
```

This installs to:

```text
~/.claude/skills/oh-my-pptx
~/.codex/skills/oh-my-pptx
```

## Use

Inspect the split pages first:

```bash
node scripts/html-to-pptx.mjs slides.html --inspect-only --keep-temp --json
```

Convert to PPTX:

```bash
node scripts/html-to-pptx.mjs slides.html --out slides.pptx
```

Use from an agent:

```text
Use $oh-my-pptx to convert this HTML slide deck into a native PPTX.
```

## Useful Options

```bash
--canvas 1920x1080
--width 1920 --height 1080
--split-mode scroll --page-count 4
--need-embed-fonts true
--script-mode keep
```

Defaults are chosen for static or agent-generated slide decks: local assets are inlined, nonessential scripts are stripped, and DeckFlow font embedding is off unless requested.

## Notes

DeckFlow is an online conversion service. Prepared HTML pages are uploaded to the configured DeckFlow backend during conversion.

Generated `.pptx` files, inspection JSON, and temporary split HTML files are intentionally ignored by Git.

## License

No license has been added yet. Treat this repository as all rights reserved until a license file is published.
