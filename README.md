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

Install or upgrade DeckFlow CLI. Oh My PPTX requires `deckops` 0.4.1 or newer:

```bash
npm install -g deckops@latest
deckops --version
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

Inspect before converting:

```bash
node scripts/html-to-pptx.mjs slides.html --inspect-only --keep-temp --json
```

Check the JSON before upload:

- `slideCount` matches the deck.
- `splitMode` is the intended mode.
- `canvas` is explicit or correctly detected.
- `warnings` are understood and fixed when they affect remote rendering.

Then convert:

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
--split-mode body
--need-embed-fonts true
--script-mode keep
--strict
```

Defaults are chosen for static or agent-generated slide decks: local assets are inlined, nonessential scripts are stripped, and DeckFlow font embedding is off unless requested. Warnings do not block conversion unless `--strict` is set.

## Splitting And Assets

`auto` first looks for slide containers such as `.slide`, `section.slide`, `[data-slide]`, and `[data-page]`. If no explicit slide containers exist, multiple `<section>` elements can be treated as slides.

For a long responsive web page, force scroll splitting and set the PPTX canvas:

```bash
node scripts/html-to-pptx.mjs page.html --inspect-only --split-mode scroll --page-count 4 --canvas 16:9 --json
node scripts/html-to-pptx.mjs page.html --out page.pptx --split-mode scroll --page-count 4 --canvas 16:9
```

Inspection warns when a generated page still contains relative asset references such as `assets/chart.png`. Those paths may work locally but fail after upload if DeckFlow receives only the generated HTML page. Inline the assets, fix missing files, or intentionally package sidecar assets before converting. Use `--no-inline-assets` or `--inline-assets false` only when that packaging has been verified.

## Troubleshooting

- Auth errors: run `deckops login`.
- Old or mismatched CLI version: upgrade with `npm install -g deckops@latest`, then confirm `deckops --version` is `0.4.1` or newer.
- Relative asset warnings: stop and fix packaging instead of retrying DeckFlow.
- Fallback canvas or one-page scroll warnings on responsive HTML: inspect the kept files, then pass `--canvas`, `--width`/`--height`, and `--split-mode scroll --page-count N`.
- JS-only, lazy-loaded, or empty pages: pre-render static HTML, eager-load media, or use `--script-mode keep` only when the deck needs JavaScript at render time.
- Use `--strict` in automated runs when any inspection warning should stop before upload.

## Notes

DeckFlow is an online conversion service. Prepared HTML pages are uploaded to the configured DeckFlow backend during conversion.

Generated `.pptx` files, inspection JSON, and temporary split HTML files are intentionally ignored by Git.

## License

No license has been added yet. Treat this repository as all rights reserved until a license file is published.
