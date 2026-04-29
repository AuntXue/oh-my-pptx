# Oh My PPTX

Oh My PPTX is an agent skill for converting HTML slide designs into native PowerPoint decks with DeckFlow's [`deckops`](https://www.npmjs.com/package/deckops) CLI.

It is designed for Claude Code, Codex, and other coding agents that generate slide layouts as HTML/CSS but need to deliver editable `.pptx` files.

## About DeckFlow

[DeckFlow](https://deckflow.com) is an AI presentation platform for creating, revamping, translating, distributing, and extending Decks. It helps teams turn documents, web pages, notes, and existing presentations into polished, brand-consistent presentation assets, and exposes stable Deck capabilities to agents and developers through interfaces such as CLI, MCP, and API.

Oh My PPTX focuses on one small part of that ecosystem: preparing HTML slide output so DeckFlow can convert it into one editable PPTX file that agents can hand back to users.

## What It Does

- Converts `.html` slide files to `.pptx` through DeckFlow/deckops.
- Splits multi-slide HTML files into one accurate single-page HTML file per slide.
- Preserves the full original DOM during selector-based splitting so `:nth-child`, `:nth-of-type`, and sibling CSS keep working.
- Supports responsive web pages through automatic `100vw`/`100vh` canvas detection or explicit scroll-slice splitting.
- Detects or assigns a PPTX-friendly canvas size, then passes width and height to DeckFlow.
- Inlines local CSS, images, and fonts by default so uploaded HTML can render without local filesystem access.
- Marks the active slide as `visible` so reveal-animation decks render their final state.
- Strips nonessential scripts from generated pages by default before upload.
- Runs a local inspection pass before upload.
- Submits all single-page HTML files to DeckFlow in one batch `html -> pptx` conversion.
- Supports DeckFlow's `--need-embed-fonts` option, defaulting to `false`.
- Downloads only the final generated PPTX for the user-facing result.
- Includes install helpers for both Claude Code and Codex skill folders.

## Requirements

- Node.js 18 or newer.
- A DeckFlow/deckops login.
- Latest DeckFlow CLI with batch HTML-to-PPTX support.
- `deckops` installed globally, or network access for the wrapper's fallback `npm exec -y deckops@latest`.

Install or upgrade deckops:

```bash
npm install -g deckops@latest
deckops login
```

You can also configure credentials manually:

```bash
deckops config set-token <token>
deckops config set-space <space-id>
```

Confirm the current CLI exposes batch conversion:

```bash
deckops convert --help
```

Look for:

```text
deckops convert <input-files...> --to pptx --width <number> --height <number> --need-embed-fonts [boolean]
```

## Install The Skill

Clone the repository:

```bash
git clone https://github.com/AuntXue/oh-my-pptx.git
cd oh-my-pptx
```

Install for both Claude Code and Codex:

```bash
./scripts/install.sh --all
```

Install for one agent runtime:

```bash
./scripts/install.sh --claude
./scripts/install.sh --codex
```

The script copies this skill to:

```text
~/.claude/skills/oh-my-pptx
~/.codex/skills/oh-my-pptx
```

## Use It Directly

Inspect first:

```bash
node scripts/html-to-pptx.mjs examples/macroverse-two-slide.html \
  --inspect-only \
  --keep-temp \
  --json
```

Convert an HTML slide file:

```bash
node scripts/html-to-pptx.mjs examples/macroverse-two-slide.html \
  --out examples/macroverse-two-slide.pptx \
  --timeout 600
```

For a multi-slide HTML file, the wrapper automatically:

- detects slide elements such as `.slide`, `[data-slide]`, and `[data-page]`
- writes temporary single-page HTML files in slide order
- keeps the full DOM and hides non-active slides for selector-based splits
- adds `visible` to the active slide for reveal-style animations
- inlines local assets
- strips scripts by default
- detects or assigns a shared canvas size
- checks the split output before upload
- converts all pages at once with `deckops convert ... --to pptx`

The wrapper writes only the final PPTX file requested with `--out`.

## Responsive Web Pages

If a slide deck uses `.slide { width: 100vw; height: 100vh; }` or `100dvh`, Oh My PPTX treats it as a responsive 16:9 deck and uses a `1920x1080` canvas by default.

When the source is a long responsive page rather than a slide deck, choose a PPTX canvas and slice the page by viewport height:

```bash
node scripts/html-to-pptx.mjs page.html \
  --split-mode scroll \
  --page-count 4 \
  --canvas 1920x1080 \
  --inspect-only \
  --keep-temp
```

After checking the temporary slide HTML files, run the same command with `--out output.pptx`.

## Scripts And Reveal Animations

Generated pages strip `<script>` tags by default. This removes navigation, editors, exporters, and other browser-only behavior that can interfere with deterministic conversion.

If content is actually created by JavaScript, keep scripts explicitly:

```bash
node scripts/html-to-pptx.mjs input.html --out output.pptx --script-mode keep
```

For reveal-animation decks, the wrapper adds `visible` to the active slide before upload, so common `.visible .reveal { opacity: 1 }` patterns render in their final state without relying on IntersectionObserver timing.

## Font Embedding

DeckFlow defaults to not embedding fonts. Oh My PPTX follows that default:

```bash
node scripts/html-to-pptx.mjs input.html --out output.pptx
```

To request embedded fonts:

```bash
node scripts/html-to-pptx.mjs input.html --out output.pptx --need-embed-fonts true
```

## Use It From An Agent

After installation, ask your agent to use the skill:

```text
Use $oh-my-pptx to convert this HTML slide deck into a native PPTX.
```

The skill instructions live in [`SKILL.md`](./SKILL.md).

## HTML Authoring Tips

For best conversion results:

- Use a fixed slide canvas, such as `1920px x 1080px`.
- Keep each slide inside a page-sized `.slide` element when possible.
- If the HTML cannot expose dimensions reliably, pass `--width <px> --height <px>` or `--canvas 16:9`.
- Prefer real HTML text for editable PowerPoint text.
- Use inline CSS or simple local stylesheets.
- Avoid hover states, JavaScript-only rendering, and lazy loading.
- Reveal animations are fine when their final state is controlled by a `visible` ancestor class.
- Include local images and fonts normally; the wrapper inlines them before upload when possible.

See [`examples/macroverse-one-slide.html`](./examples/macroverse-one-slide.html) for a minimal one-slide example, [`examples/macroverse-two-slide.html`](./examples/macroverse-two-slide.html) for a multi-slide example, and [`examples/responsive-scroll-page.html`](./examples/responsive-scroll-page.html) for scroll splitting.

## Repository Layout

```text
oh-my-pptx/
├── SKILL.md
├── agents/openai.yaml
├── assets/
│   ├── deckflow-icon-400.png
│   └── deckflow-icon-1024.png
├── examples/macroverse-one-slide.html
├── examples/macroverse-two-slide.html
├── examples/responsive-scroll-page.html
└── scripts/
    ├── html-to-pptx.mjs
    └── install.sh
```

## Notes

Deckops is an online conversion tool. Prepared HTML files are uploaded to the configured DeckFlow backend during conversion.

Generated files such as `.pptx`, inspection JSON, and temporary split HTML files are intentionally ignored by Git.

## License

No license has been added yet. Treat this repository as all rights reserved until a license file is published.
