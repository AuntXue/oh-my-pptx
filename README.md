# Oh My PPTX

Oh My PPTX is an agent skill for converting HTML slide designs into native PowerPoint decks with DeckFlow's [`deckops`](https://www.npmjs.com/package/deckops) CLI.

It is designed for Claude Code, Codex, and other coding agents that generate slide layouts as HTML/CSS but need to deliver editable `.pptx` files.

## About DeckFlow

[DeckFlow](https://deckflow.com) is an AI presentation platform for creating, revamping, translating, distributing, and extending Decks. It helps teams turn documents, web pages, notes, and existing presentations into polished, brand-consistent presentation assets, and exposes stable Deck capabilities to agents and developers through interfaces such as CLI, MCP, and API.

Oh My PPTX focuses on one small part of that ecosystem: using DeckFlow/deckops to turn HTML slide output into one editable PPTX file that agents can hand back to users.

## What It Does

- Converts `.html` slide files to `.pptx` through DeckFlow/deckops.
- Splits multi-slide HTML files into one slide per DeckFlow conversion request.
- Detects slide width and height, then passes those dimensions to DeckFlow for accurate PPTX sizing.
- Joins multiple one-slide PPTX files back into one final deck with `deckops join`.
- Wraps the exact `deckops` command in a small Node.js script.
- Downloads only the final generated PPTX for the user-facing result.
- Includes install helpers for both Claude Code and Codex skill folders.

## Requirements

- Node.js 18 or newer.
- A Deckops/DeckFlow login.
- A join-enabled DeckFlow CLI. The wrapper uses `deckops join` for multi-page output.
- `deckops` installed globally, or network access for the wrapper's fallback `npm exec -y deckops@latest`.

Install and authenticate deckops:

```bash
npm install -g deckops
deckops login
```

You can also configure credentials manually:

```bash
deckops config set-token <token>
deckops config set-space <space-id>
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

Convert an HTML slide file:

```bash
node scripts/html-to-pptx.mjs examples/macroverse-one-slide.html \
  --out examples/macroverse-one-slide.pptx \
  --timeout 300
```

For a multi-slide HTML file, the wrapper automatically:

- detects each `.slide`
- writes temporary single-slide HTML files
- detects width and height for each slide
- converts each slide with `deckops run convertor.html2pptx ... --param width=<px> --param height=<px>`
- joins the intermediate PPTX files with `deckops join`

The wrapper writes only the final PPTX file requested with `--out`.

## Use It From An Agent

After installation, ask your agent to use the skill:

```text
Use $oh-my-pptx to convert this HTML slide into a native PPTX deck.
```

The skill instructions live in [`SKILL.md`](./SKILL.md).

## HTML Authoring Tips

For best conversion results:

- Use a fixed slide canvas, such as `1920px x 1080px`.
- Keep each slide inside a page-sized `.slide` element.
- If the HTML cannot expose dimensions reliably, pass `--width <px> --height <px>`.
- Prefer real HTML text for editable PowerPoint text.
- Use inline CSS or simple local stylesheets.
- Avoid animations, hover states, JavaScript-only rendering, and viewport-dependent layouts.
- Make images and fonts available to the conversion environment.

See [`examples/macroverse-one-slide.html`](./examples/macroverse-one-slide.html) for a minimal one-slide example and [`examples/macroverse-two-slide.html`](./examples/macroverse-two-slide.html) for a multi-slide example.

## Repository Layout

```text
oh-my-pptx/
├── SKILL.md
├── agents/openai.yaml
├── examples/macroverse-one-slide.html
├── examples/macroverse-two-slide.html
└── scripts/
    ├── html-to-pptx.mjs
    └── install.sh
```

## Notes

Deckops is an online conversion tool. HTML files are uploaded to the configured DeckFlow backend during conversion.

Generated files such as `.pptx`, `.task.json`, and extracted preview images are intentionally ignored by Git.

## License

No license has been added yet. Treat this repository as all rights reserved until a license file is published.
