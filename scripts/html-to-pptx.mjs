#!/usr/bin/env node

import { spawn, spawnSync } from "node:child_process";
import {
  createWriteStream,
  existsSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  rmSync,
  statSync,
  writeFileSync,
} from "node:fs";
import { tmpdir } from "node:os";
import { dirname, extname, isAbsolute, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const DEFAULT_WIDTH = 1920;
const DEFAULT_HEIGHT = 1080;
const DEFAULT_SELECTORS = [".slide", "section.slide", "[data-slide]", "[data-slide-index]", "[data-page]"];
const VOID_TAGS = new Set([
  "area",
  "base",
  "br",
  "col",
  "embed",
  "hr",
  "img",
  "input",
  "link",
  "meta",
  "source",
  "track",
  "wbr",
]);

const USAGE = `
Usage:
  html-to-pptx.mjs <input.html> --out output.pptx [options]

Options:
  --out <file>                   Final PPTX path. Required unless --inspect-only is used.
  --timeout <sec>                DeckFlow task timeout. Default: 600.
  --deckops <cmd>                Deckops command. Default: DECKOPS_BIN, deckops, or npm exec fallback.
  --split-mode <mode>            auto, selector, scroll, or body. Default: auto.
  --slide-selector <selectors>   Comma-separated simple selectors. Default: common slide selectors.
  --canvas <size>                auto, 16:9, 4:3, or WIDTHxHEIGHT. Default: auto.
  --width <px>                   Override canvas width.
  --height <px>                  Override canvas height.
  --page-count <n>               Page count for scroll splitting when no slide elements exist.
  --need-embed-fonts [boolean]   Pass DeckFlow needEmbedFonts. Default: false.
  --inspect-only                 Split and validate HTML pages without uploading to DeckFlow.
  --strict                       Treat inspection warnings as blocking errors.
  --inline-assets <boolean>      Inline local CSS/images/fonts as data URIs. Default: true.
  --script-mode <mode>           strip or keep scripts in generated pages. Default: strip.
  --keep-temp                    Keep temporary single-page HTML and inspection files.
  --json                         Print machine-readable summary. Does not write task JSON files.
  --help                         Show this help.

Environment:
  DECKOPS_BIN, DECKOPS_NPM_SPEC
`;

function parseArgs(argv) {
  const args = {
    input: null,
    out: null,
    timeout: "600",
    deckops: process.env.DECKOPS_BIN || null,
    splitMode: "auto",
    slideSelector: null,
    canvas: "auto",
    width: null,
    height: null,
    pageCount: null,
    needEmbedFonts: false,
    inspectOnly: false,
    strict: false,
    inlineAssets: true,
    scriptMode: "strip",
    keepTemp: false,
    json: false,
  };

  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    if (arg === "--help" || arg === "-h") {
      console.log(USAGE.trim());
      process.exit(0);
    } else if (arg === "--out") {
      args.out = argv[++i];
    } else if (arg === "--timeout") {
      args.timeout = argv[++i];
    } else if (arg === "--deckops") {
      args.deckops = argv[++i];
    } else if (arg === "--split-mode") {
      args.splitMode = argv[++i];
    } else if (arg === "--slide-selector") {
      args.slideSelector = argv[++i];
    } else if (arg === "--canvas") {
      args.canvas = argv[++i];
    } else if (arg === "--width") {
      args.width = parsePositiveInt(argv[++i], "--width");
    } else if (arg === "--height") {
      args.height = parsePositiveInt(argv[++i], "--height");
    } else if (arg === "--page-count") {
      args.pageCount = parsePositiveInt(argv[++i], "--page-count");
    } else if (arg === "--need-embed-fonts") {
      args.needEmbedFonts = parseOptionalBoolean(argv[i + 1], true);
      if (isBooleanLiteral(argv[i + 1])) i += 1;
    } else if (arg === "--inspect-only") {
      args.inspectOnly = true;
    } else if (arg === "--strict") {
      args.strict = true;
    } else if (arg === "--inline-assets") {
      args.inlineAssets = parseRequiredBoolean(argv[++i], "--inline-assets");
    } else if (arg === "--no-inline-assets") {
      args.inlineAssets = false;
    } else if (arg === "--script-mode") {
      args.scriptMode = argv[++i];
    } else if (arg === "--keep-temp") {
      args.keepTemp = true;
    } else if (arg === "--json") {
      args.json = true;
    } else if (arg === "--width-param" || arg === "--height-param" || arg === "--convert-task" || arg === "--param") {
      i += 1;
    } else if (!args.input) {
      args.input = arg;
    } else {
      throw new Error(`Unexpected argument: ${arg}`);
    }
  }

  if (!args.input) throw new Error("Missing input HTML file.");
  if (!args.out && !args.inspectOnly) throw new Error("Missing required --out output.pptx path.");
  if (args.out && extname(args.out).toLowerCase() !== ".pptx") throw new Error("--out must end with .pptx.");
  if (!args.timeout || !/^\d+$/.test(args.timeout)) {
    throw new Error("--timeout must be a positive integer.");
  }
  if (!["auto", "selector", "scroll", "body"].includes(args.splitMode)) {
    throw new Error("--split-mode must be auto, selector, scroll, or body.");
  }
  if (!["strip", "keep"].includes(args.scriptMode)) {
    throw new Error("--script-mode must be strip or keep.");
  }
  parseCanvas(args.canvas);
  return args;
}

function parsePositiveInt(value, name) {
  const number = Number(value);
  if (!Number.isInteger(number) || number <= 0) {
    throw new Error(`${name} must be a positive integer.`);
  }
  return number;
}

function parseOptionalBoolean(value, defaultValue) {
  if (!isBooleanLiteral(value)) return defaultValue;
  return value.toLowerCase() === "true";
}

function parseRequiredBoolean(value, name) {
  if (!isBooleanLiteral(value)) throw new Error(`${name} must be true or false.`);
  return value.toLowerCase() === "true";
}

function isBooleanLiteral(value) {
  return typeof value === "string" && /^(true|false)$/i.test(value);
}

function commandExists(cmd) {
  const result = spawnSync(cmd, ["--help"], { stdio: "ignore" });
  return result.status === 0 || result.status === 1;
}

function deckopsCommand(explicit) {
  if (explicit) return { cmd: explicit, baseArgs: [] };
  if (commandExists("deckops")) return { cmd: "deckops", baseArgs: [] };
  return {
    cmd: "npm",
    baseArgs: ["exec", "-y", process.env.DECKOPS_NPM_SPEC || "deckops@latest", "--"],
  };
}

function runDeckops(command, deckopsArgs) {
  const fullArgs = [...command.baseArgs, "--json", ...deckopsArgs];
  return new Promise((resolvePromise, reject) => {
    const child = spawn(command.cmd, fullArgs, { stdio: ["ignore", "pipe", "pipe"] });
    let stdout = "";
    let stderr = "";

    child.stdout.on("data", (chunk) => {
      stdout += chunk.toString();
    });
    child.stderr.on("data", (chunk) => {
      stderr += chunk.toString();
      process.stderr.write(chunk);
    });
    child.on("error", reject);
    child.on("close", (code) => {
      if (code !== 0) {
        reject(new Error(`deckops exited with code ${code}${stderr ? `\n${stderr}` : ""}`));
      } else {
        resolvePromise(stdout);
      }
    });
  });
}

function parseJsonOutput(stdout) {
  const clean = stripAnsi(stdout).trim();
  try {
    return JSON.parse(clean);
  } catch {
    const objectStart = clean.indexOf("{");
    const objectEnd = clean.lastIndexOf("}");
    if (objectStart >= 0 && objectEnd > objectStart) return JSON.parse(clean.slice(objectStart, objectEnd + 1));
    const arrayStart = clean.indexOf("[");
    const arrayEnd = clean.lastIndexOf("]");
    if (arrayStart >= 0 && arrayEnd > arrayStart) return JSON.parse(clean.slice(arrayStart, arrayEnd + 1));
    const task = parseHumanTaskOutput(clean);
    if (task) return task;
    throw new Error(`Could not parse deckops output:\n${stdout}`);
  }
}

function stripAnsi(value) {
  return value.replace(/\x1B\[[0-?]*[ -/]*[@-~]/g, "");
}

function parseHumanTaskOutput(output) {
  const resultStart = output.indexOf("Result:");
  if (resultStart < 0) return null;
  const resultText = output.slice(resultStart + "Result:".length).trim();
  const arrayStart = resultText.indexOf("[");
  const arrayEnd = resultText.lastIndexOf("]");
  if (arrayStart < 0 || arrayEnd <= arrayStart) return null;
  let result;
  try {
    result = JSON.parse(resultText.slice(arrayStart, arrayEnd + 1));
  } catch {
    return null;
  }
  return {
    id: output.match(/Task ID:\s*(\S+)/)?.[1],
    type: output.match(/Type:\s*(\S+)/)?.[1],
    status: output.match(/Status:\s*(\S+)/)?.[1],
    result,
  };
}

function collectUrls(value, urls = []) {
  if (!value) return urls;
  if (typeof value === "string") {
    if (/^https?:\/\//i.test(value)) urls.push(value);
    return urls;
  }
  if (Array.isArray(value)) {
    for (const item of value) collectUrls(item, urls);
    return urls;
  }
  if (typeof value === "object") {
    for (const item of Object.values(value)) collectUrls(item, urls);
  }
  return urls;
}

function preferPptxUrl(urls) {
  return (
    urls.find((url) => /\.pptx(?:[?#]|$)/i.test(url)) ||
    urls.find((url) => /pptx/i.test(url)) ||
    urls[0] ||
    null
  );
}

async function download(url, outPath) {
  const response = await fetch(url);
  if (!response.ok) throw new Error(`Download failed: ${response.status} ${response.statusText}`);
  mkdirSync(dirname(outPath), { recursive: true });
  const file = createWriteStream(outPath);
  await new Promise((resolvePromise, reject) => {
    response.body
      .pipeTo(
        new WritableStream({
          write(chunk) {
            file.write(Buffer.from(chunk));
          },
          close() {
            file.end(resolvePromise);
          },
          abort(reason) {
            file.destroy(reason);
            reject(reason);
          },
        })
      )
      .catch(reject);
  });
}

function parseCanvas(value) {
  if (!value || value === "auto") return null;
  if (value === "16:9") return { width: DEFAULT_WIDTH, height: DEFAULT_HEIGHT, source: "canvas 16:9" };
  if (value === "4:3") return { width: 1600, height: 1200, source: "canvas 4:3" };
  const match = value.match(/^(\d+)x(\d+)$/i);
  if (!match) throw new Error("--canvas must be auto, 16:9, 4:3, or WIDTHxHEIGHT.");
  return { width: Number(match[1]), height: Number(match[2]), source: "canvas option" };
}

function selectorList(args) {
  if (!args.slideSelector) return DEFAULT_SELECTORS;
  return args.slideSelector
    .split(",")
    .map((item) => item.trim())
    .filter(Boolean);
}

function prepareHtml(rawHtml, inputDir, args, globalWarnings) {
  if (!args.inlineAssets) return rawHtml;
  return inlineAssets(rawHtml, inputDir, globalWarnings);
}

function inlineAssets(html, inputDir, warnings) {
  let result = html.replace(/<link\b([^>]*?)>/gi, (tag, attrs) => {
    const rel = attrValue(attrs, "rel");
    const href = attrValue(attrs, "href");
    if (!rel || !/\bstylesheet\b/i.test(rel) || !href) return tag;
    const assetPath = resolveAssetPath(href, inputDir, warnings);
    if (!assetPath) return tag;
    const css = readFileSync(assetPath, "utf8");
    const inlined = inlineCssUrls(css, dirname(assetPath), warnings);
    return `<style data-oh-my-pptx-inlined-from="${escapeHtmlAttr(href)}">\n${inlined}\n</style>`;
  });

  result = result.replace(/<style\b([^>]*)>([\s\S]*?)<\/style>/gi, (tag, attrs, css) => {
    return `<style${attrs}>${inlineCssUrls(css, inputDir, warnings)}</style>`;
  });

  result = result.replace(/\bstyle\s*=\s*("[^"]*"|'[^']*')/gi, (match, quoted) => {
    const quote = quoted[0];
    const style = quoted.slice(1, -1);
    return `style=${quote}${inlineCssUrls(style, inputDir, warnings)}${quote}`;
  });

  result = result.replace(/<([a-zA-Z][\w:-]*)(\s[^>]*)?>/g, (tag, tagName, attrs = "") => {
    const lower = tagName.toLowerCase();
    const assetAttrs = lower === "image" ? ["href", "xlink:href"] : ["src", "poster"];
    if (!["img", "image", "source", "video", "audio", "embed"].includes(lower)) return tag;
    let next = tag;
    for (const name of assetAttrs) {
      next = replaceAssetAttribute(next, name, inputDir, warnings);
    }
    return next;
  });

  return result;
}

function replaceAssetAttribute(tag, name, inputDir, warnings) {
  const attrRegex = new RegExp(`\\b${escapeRegExp(name)}\\s*=\\s*("[^"]*"|'[^']*'|[^\\s>]+)`, "i");
  return tag.replace(attrRegex, (match, quoted) => {
    const raw = quoted.replace(/^['"]|['"]$/g, "");
    const assetPath = resolveAssetPath(raw, inputDir, warnings);
    if (!assetPath) return match;
    return `${name}="${toDataUri(assetPath)}"`;
  });
}

function inlineCssUrls(css, baseDir, warnings) {
  return css.replace(/url\(\s*(['"]?)([^'")]+)\1\s*\)/gi, (match, quote, rawUrl) => {
    const assetPath = resolveAssetPath(rawUrl, baseDir, warnings);
    if (!assetPath) return match;
    return `url("${toDataUri(assetPath)}")`;
  });
}

function resolveAssetPath(rawUrl, baseDir, warnings) {
  const cleaned = String(rawUrl).trim().replace(/^['"]|['"]$/g, "");
  if (!cleaned || /^(?:data|https?|blob|mailto|javascript):/i.test(cleaned) || cleaned.startsWith("#")) return null;
  const withoutQuery = cleaned.split(/[?#]/)[0];
  let filePath;
  try {
    if (/^file:/i.test(withoutQuery)) {
      filePath = fileURLToPath(withoutQuery);
    } else if (isAbsolute(withoutQuery)) {
      filePath = withoutQuery;
    } else {
      filePath = resolve(baseDir, withoutQuery);
    }
  } catch {
    warnings.push(`Could not resolve asset URL: ${cleaned}`);
    return null;
  }

  if (!existsSync(filePath)) {
    warnings.push(`Missing local asset: ${cleaned}`);
    return null;
  }
  if (!statSync(filePath).isFile()) {
    warnings.push(`Asset is not a file: ${cleaned}`);
    return null;
  }
  return filePath;
}

function toDataUri(filePath) {
  return `data:${mimeType(filePath)};base64,${readFileSync(filePath).toString("base64")}`;
}

function mimeType(filePath) {
  const ext = extname(filePath).toLowerCase();
  const types = {
    ".apng": "image/apng",
    ".avif": "image/avif",
    ".css": "text/css",
    ".gif": "image/gif",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".otf": "font/otf",
    ".png": "image/png",
    ".svg": "image/svg+xml",
    ".ttf": "font/ttf",
    ".webp": "image/webp",
    ".woff": "font/woff",
    ".woff2": "font/woff2",
  };
  return types[ext] || "application/octet-stream";
}

function attrValue(attrs, name) {
  const match = attrs.match(new RegExp(`\\b${escapeRegExp(name)}\\s*=\\s*("[^"]*"|'[^']*'|[^\\s>]+)`, "i"));
  return match ? match[1].replace(/^['"]|['"]$/g, "") : null;
}

function createPages(html, args, inputPath, globalWarnings) {
  const inputDir = dirname(inputPath);
  const selectors = selectorList(args);
  const canvas = detectCanvas(html, selectors, args);
  const ranges = args.splitMode === "scroll" || args.splitMode === "body"
    ? []
    : findSlideRanges(html, selectors, args);

  if (ranges.length > 0) {
    return {
      canvas,
      splitMode: "selector",
      selectors,
      pages: ranges.map((range, index) => ({
        index: index + 1,
        html: finalizePageHtml(buildSelectorPageHtml(html, ranges, index, canvas), args),
        source: `${range.selector} ${index + 1}`,
      })),
    };
  }

  if (args.splitMode === "selector") {
    throw new Error(`No slide elements matched selectors: ${selectors.join(", ")}`);
  }

  if (args.splitMode === "body") {
    return {
      canvas,
      splitMode: "body",
      selectors,
      pages: [{
        index: 1,
        html: finalizePageHtml(injectRuntimeStyle(html, canvas, "body", 0), args),
        source: "body",
      }],
    };
  }

  const pageCount = args.pageCount || detectDeclaredPageCount(html) || 1;
  if (!args.pageCount && pageCount === 1) {
    globalWarnings.push("No slide elements found; using one scroll slice. Pass --split-mode scroll --page-count N for long responsive pages.");
  }

  return {
    canvas,
    splitMode: "scroll",
    selectors,
    pages: Array.from({ length: pageCount }, (_, index) => ({
      index: index + 1,
      html: finalizePageHtml(injectRuntimeStyle(html, canvas, "scroll", index * canvas.height), args),
      source: `scroll ${index + 1}`,
    })),
  };
}

function finalizePageHtml(html, args) {
  return args.scriptMode === "strip" ? stripScripts(html) : html;
}

function stripScripts(html) {
  return html.replace(/<script\b[\s\S]*?<\/script>/gi, "");
}

function detectCanvas(html, selectors, args) {
  if (args.width && args.height) return { width: args.width, height: args.height, source: "override" };
  const canvasOption = parseCanvas(args.canvas);
  if (canvasOption) {
    return {
      width: args.width || canvasOption.width,
      height: args.height || canvasOption.height,
      source: canvasOption.source,
    };
  }

  const firstRange = findSlideRanges(html, selectors, { ...args, splitMode: "auto" })[0];
  const slideHtml = firstRange ? html.slice(firstRange.start, firstRange.end) : "";
  const detected =
    mergeSize(args, detectDataSize(slideHtml)) ||
    mergeSize(args, detectInlineStyleSize(slideHtml)) ||
    mergeSize(args, detectCssRuleSize(html, selectors)) ||
    mergeSize(args, detectPageSize(html)) ||
    mergeSize(args, detectResponsiveViewportSize(html, selectors)) ||
    mergeSize(args, detectViewportSize(html));

  if (detected) return detected;
  return { width: args.width || DEFAULT_WIDTH, height: args.height || DEFAULT_HEIGHT, source: "fallback 16:9" };
}

function mergeSize(args, detected) {
  if (!detected) return null;
  const width = args.width || detected.width;
  const height = args.height || detected.height;
  if (width && height) return { width, height, source: detected.source };
  return null;
}

function detectDataSize(fragment) {
  const open = fragment.match(/<[^>]+>/);
  if (!open) return null;
  const width = attrNumber(open[0], "data-width") || attrNumber(open[0], "width");
  const height = attrNumber(open[0], "data-height") || attrNumber(open[0], "height");
  return width && height ? { width, height, source: "slide attributes" } : null;
}

function attrNumber(tag, name) {
  const match = tag.match(new RegExp(`\\b${escapeRegExp(name)}\\s*=\\s*("[^"]*"|'[^']*'|[^\\s>]+)`, "i"));
  if (!match) return null;
  return parseCssPixelValue(match[1].replace(/^['"]|['"]$/g, ""));
}

function detectInlineStyleSize(fragment) {
  const open = fragment.match(/<[^>]+>/);
  if (!open) return null;
  const styleMatch = open[0].match(/\bstyle\s*=\s*("[^"]*"|'[^']*')/i);
  if (!styleMatch) return null;
  const style = styleMatch[1].slice(1, -1);
  const width = declarationPixel(style, "width");
  const height = declarationPixel(style, "height");
  return width && height ? { width, height, source: "slide inline style" } : null;
}

function detectCssRuleSize(html, selectors) {
  const styleBlocks = [...html.matchAll(/<style\b[^>]*>([\s\S]*?)<\/style>/gi)].map((m) => m[1]);
  for (const css of styleBlocks) {
    const ruleRegex = /([^{}]+)\{([^{}]+)\}/g;
    let match;
    while ((match = ruleRegex.exec(css))) {
      const ruleSelectors = match[1].split(",").map((item) => item.trim());
      if (!ruleSelectors.some((ruleSelector) => selectors.some((selector) => selectorLooksCompatible(ruleSelector, selector)))) continue;
      const width = declarationPixel(match[2], "width");
      const height = declarationPixel(match[2], "height");
      if (width && height) return { width, height, source: `${ruleSelectors[0]} CSS rule` };
    }
  }
  return null;
}

function detectResponsiveViewportSize(html, selectors) {
  const styleBlocks = [...html.matchAll(/<style\b[^>]*>([\s\S]*?)<\/style>/gi)].map((m) => m[1]);
  for (const css of styleBlocks) {
    const ruleRegex = /([^{}]+)\{([^{}]+)\}/g;
    let match;
    while ((match = ruleRegex.exec(css))) {
      const ruleSelectors = match[1].split(",").map((item) => item.trim());
      if (!ruleSelectors.some((ruleSelector) => selectors.some((selector) => selectorLooksCompatible(ruleSelector, selector)))) continue;
      if (hasViewportWidth(match[2]) && hasViewportHeight(match[2])) {
        return { width: DEFAULT_WIDTH, height: DEFAULT_HEIGHT, source: "responsive viewport 16:9" };
      }
    }
  }
  return null;
}

function hasViewportWidth(css) {
  return /(?:^|;)\s*width\s*:\s*100(?:d?vw|%)/i.test(css);
}

function hasViewportHeight(css) {
  return /(?:^|;)\s*height\s*:\s*100d?vh/i.test(css) || /(?:^|;)\s*min-height\s*:\s*100d?vh/i.test(css);
}

function selectorLooksCompatible(ruleSelector, selector) {
  if (selector.startsWith(".") && selectorMatchesClass(ruleSelector, selector.slice(1))) return true;
  if (selector.startsWith("[") && ruleSelector.includes(selector)) return true;
  return ruleSelector === selector;
}

function selectorMatchesClass(selector, className) {
  return new RegExp(`(^|[^\\w-])\\.${escapeRegExp(className)}($|[^\\w-])`).test(selector);
}

function detectPageSize(html) {
  const match = html.match(/@page\s*(?:[^{]*)\{[^}]*\bsize\s*:\s*(\d+(?:\.\d+)?)px\s+(\d+(?:\.\d+)?)px/i);
  return match
    ? { width: Math.round(Number(match[1])), height: Math.round(Number(match[2])), source: "@page size" }
    : null;
}

function detectViewportSize(html) {
  const meta = html.match(/<meta\b[^>]*name\s*=\s*["']viewport["'][^>]*>/i);
  if (!meta) return null;
  const contentMatch = meta[0].match(/\bcontent\s*=\s*("[^"]*"|'[^']*')/i);
  if (!contentMatch) return null;
  const content = contentMatch[1].slice(1, -1);
  const width = content.match(/(?:^|,)\s*width\s*=\s*(\d+)/i);
  const height = content.match(/(?:^|,)\s*height\s*=\s*(\d+)/i);
  return width && height
    ? { width: Number(width[1]), height: Number(height[1]), source: "viewport meta" }
    : null;
}

function declarationPixel(css, property) {
  const match = css.match(new RegExp(`(?:^|;)\\s*${property}\\s*:\\s*([^;]+)`, "i"));
  return match ? parseCssPixelValue(match[1]) : null;
}

function parseCssPixelValue(value) {
  const cleaned = String(value).trim().toLowerCase();
  const match = cleaned.match(/^(\d+(?:\.\d+)?)(px)?$/);
  return match ? Math.round(Number(match[1])) : null;
}

function detectDeclaredPageCount(html) {
  const meta = html.match(/<meta\b[^>]*name\s*=\s*["']oh-my-pptx-pages["'][^>]*>/i);
  const content = meta?.[0]?.match(/\bcontent\s*=\s*("(\d+)"|'(\d+)')/i);
  if (content) return Number(content[2] || content[3]);
  const dataPages = html.match(/\bdata-pages\s*=\s*["']?(\d+)/i);
  return dataPages ? Number(dataPages[1]) : null;
}

function findSlideRanges(html, selectors, args) {
  const ranges = findRangesBySelectors(html, selectors);
  if (ranges.length > 0) return ranges;
  if (args.slideSelector || args.splitMode === "selector") return [];

  const sectionRanges = findRangesBySelectors(html, ["section"]);
  return sectionRanges.length > 1 ? sectionRanges : [];
}

function findRangesBySelectors(html, selectors) {
  const ranges = [];
  const tagRegex = /<([a-zA-Z][\w:-]*)(\s[^>]*)?>/g;
  let match;
  while ((match = tagRegex.exec(html))) {
    const [tagText, tagName, attrsText = ""] = match;
    if (isVoidOrSelfClosing(tagName, tagText)) continue;
    const attrs = parseAttributes(attrsText);
    const selector = selectors.find((candidate) => matchesSimpleSelector(tagName, attrs, candidate));
    if (!selector) continue;
    const end = findMatchingClose(html, tagName, tagRegex.lastIndex);
    if (end) {
      ranges.push({
        start: match.index,
        openEnd: tagRegex.lastIndex,
        end,
        tagName,
        selector,
      });
    }
  }
  return dedupeAndRemoveNested(ranges);
}

function parseAttributes(attrsText) {
  const attrs = {};
  const attrRegex = /\s+([\w:-]+)(?:\s*=\s*("[^"]*"|'[^']*'|[^\s"'=<>`]+))?/g;
  let match;
  while ((match = attrRegex.exec(attrsText))) {
    attrs[match[1].toLowerCase()] = match[2] ? match[2].replace(/^['"]|['"]$/g, "") : "";
  }
  return attrs;
}

function matchesSimpleSelector(tagName, attrs, selector) {
  const trimmed = selector.trim();
  if (!trimmed || /[\s>+~]/.test(trimmed)) return false;
  let rest = trimmed;

  const attrMatches = [...rest.matchAll(/\[([\w:-]+)(?:\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\]]+)))?\]/g)];
  for (const match of attrMatches) {
    const name = match[1].toLowerCase();
    const expected = match[2] ?? match[3] ?? match[4];
    if (!(name in attrs)) return false;
    if (expected !== undefined && attrs[name] !== String(expected).replace(/^['"]|['"]$/g, "")) return false;
  }
  rest = rest.replace(/\[[^\]]+\]/g, "");

  const classMatches = [...rest.matchAll(/\.([\w-]+)/g)];
  const classes = (attrs.class || "").split(/\s+/).filter(Boolean);
  for (const match of classMatches) {
    if (!classes.includes(match[1])) return false;
  }
  rest = rest.replace(/\.[\w-]+/g, "").trim();

  return !rest || rest === "*" || rest.toLowerCase() === tagName.toLowerCase();
}

function dedupeAndRemoveNested(ranges) {
  const unique = [...new Map(ranges.map((range) => [range.start, range])).values()].sort((a, b) => a.start - b.start);
  return unique.filter((range) => !unique.some((other) => other !== range && other.start < range.start && other.end >= range.end));
}

function isVoidOrSelfClosing(tagName, tagText) {
  return tagText.endsWith("/>") || VOID_TAGS.has(tagName.toLowerCase());
}

function findMatchingClose(html, tagName, fromIndex) {
  const escaped = escapeRegExp(tagName);
  const tagRegex = new RegExp(`<\\/?${escaped}(?:\\s[^>]*)?>`, "gi");
  tagRegex.lastIndex = fromIndex;
  let depth = 1;
  let match;
  while ((match = tagRegex.exec(html))) {
    const tagText = match[0];
    if (tagText.startsWith("</")) {
      depth -= 1;
      if (depth === 0) return tagRegex.lastIndex;
    } else if (!isVoidOrSelfClosing(tagName, tagText)) {
      depth += 1;
    }
  }
  return null;
}

function buildSelectorPageHtml(html, ranges, activeIndex, canvas) {
  let output = "";
  let cursor = 0;
  for (let index = 0; index < ranges.length; index += 1) {
    const range = ranges[index];
    output += html.slice(cursor, range.start);
    const openTag = html.slice(range.start, range.openEnd);
    output += markSlideOpenTag(openTag, index + 1, index === activeIndex);
    cursor = range.openEnd;
  }
  output += html.slice(cursor);
  return injectRuntimeStyle(output, canvas, "selector", 0);
}

function markSlideOpenTag(openTag, pageNumber, active) {
  let next = openTag.replace(/>$/, ` data-oh-my-pptx-page="${pageNumber}"${active ? " data-oh-my-pptx-active=\"1\"" : " data-oh-my-pptx-hidden=\"1\""}>`);
  if (!active) return next;

  if (/\bclass\s*=/.test(next)) {
    return next.replace(/\bclass\s*=\s*("[^"]*"|'[^']*'|[^\s>]+)/i, (match, quoted) => {
      const quote = quoted[0] === "'" ? "'" : "\"";
      const raw = quoted.replace(/^['"]|['"]$/g, "");
      const classes = raw.split(/\s+/).filter(Boolean);
      if (!classes.includes("visible")) classes.push("visible");
      return `class=${quote}${classes.join(" ")}${quote}`;
    });
  }
  return next.replace(/>$/, ` class="visible">`);
}

function injectRuntimeStyle(html, canvas, mode, offsetY) {
  const modeCss = mode === "selector"
    ? `
[data-oh-my-pptx-hidden="1"] { display: none !important; }
[data-oh-my-pptx-active="1"] {
  max-width: none !important;
  max-height: none !important;
  overflow: hidden !important;
}`
    : mode === "scroll"
      ? `
body {
  width: ${canvas.width}px !important;
  min-height: ${canvas.height}px !important;
  overflow: visible !important;
  transform: translateY(-${offsetY}px);
  transform-origin: 0 0;
}`
      : "";

  const frameCss = mode === "selector"
    ? `
html,
body {
  margin: 0 !important;
  overflow: hidden !important;
}`
    : `
html {
  margin: 0 !important;
  width: ${canvas.width}px !important;
  height: ${canvas.height}px !important;
  overflow: hidden !important;
}
body {
  margin: 0 !important;
  width: ${canvas.width}px !important;
  height: ${canvas.height}px !important;
  overflow: hidden !important;
}`;

  const css = `
<style id="oh-my-pptx-runtime">
${frameCss}
${modeCss}
</style>`;
  return injectIntoHead(html, css);
}

function injectIntoHead(html, content) {
  if (/<\/head>/i.test(html)) return html.replace(/<\/head>/i, `${content}\n</head>`);
  if (/<html\b[^>]*>/i.test(html)) return html.replace(/<html\b[^>]*>/i, (match) => `${match}\n<head><meta charset="utf-8">${content}</head>`);
  return `<!doctype html><html><head><meta charset="utf-8">${content}</head><body>${html}</body></html>`;
}

function inspectPages(pages, canvas, splitMode, scriptMode, globalWarnings) {
  const errors = [];
  const warnings = [...globalWarnings];

  if (pages.length === 0) errors.push("No pages were generated.");
  if (!canvas.width || !canvas.height) errors.push("Canvas width/height could not be determined.");
  if (canvas.width / canvas.height < 1 || canvas.width / canvas.height > 2.5) {
    warnings.push(`Canvas aspect ratio looks unusual: ${canvas.width}x${canvas.height}.`);
  }
  if (canvas.source.startsWith("fallback")) {
    warnings.push("Canvas size came from fallback 1920x1080; pass --width/--height or --canvas if the source is responsive.");
  }
  if (splitMode === "scroll" && pages.length === 1) {
    warnings.push("Scroll split produced one page. For long responsive pages, pass --page-count after inspecting rendered height.");
  }

  const pageReports = pages.map((page) => {
    const pageWarnings = [];
    const visibleHtml = activeSlideHtml(page.html) || page.html;
    const text = stripTags(visibleHtml).replace(/\s+/g, " ").trim();
    if (text.length === 0) pageWarnings.push("Page has no visible text in raw HTML; verify it is not JS-only or image-only.");
    if (/<script\b/i.test(page.html)) pageWarnings.push("Page contains scripts; DeckFlow rendering may not wait for dynamic content.");
    if (/\bloading\s*=\s*["']lazy["']/i.test(visibleHtml)) pageWarnings.push("Page contains lazy-loaded media; inline or eager-load it before conversion.");

    const relativeAssets = findRelativeAssets(page.html);
    if (relativeAssets.length > 0) {
      pageWarnings.push(`Page still has relative asset references: ${relativeAssets.slice(0, 5).join(", ")}`);
      pageWarnings.push("Relative asset references can fail in DeckFlow if the uploaded HTML page does not carry those files with it; inline assets or repackage the slide before conversion.");
    }

    warnings.push(...pageWarnings.map((warning) => `slide ${page.index}: ${warning}`));
    return {
      slide: page.index,
      source: page.source,
      bytes: Buffer.byteLength(page.html, "utf8"),
      textLength: text.length,
      warnings: pageWarnings,
    };
  });

  return {
    ok: errors.length === 0,
    canvas,
    splitMode,
    scriptMode,
    slideCount: pages.length,
    pages: pageReports,
    warnings,
    errors,
  };
}

function activeSlideHtml(html) {
  const activeOpen = /<([a-zA-Z][\w:-]*)(?=[^>]*\bdata-oh-my-pptx-active\s*=\s*["']1["'])[^>]*>/i.exec(html);
  if (!activeOpen) return null;
  const end = findMatchingClose(html, activeOpen[1], activeOpen.index + activeOpen[0].length);
  return end ? html.slice(activeOpen.index, end) : activeOpen[0];
}

function stripTags(html) {
  return html
    .replace(/<script\b[\s\S]*?<\/script>/gi, " ")
    .replace(/<style\b[\s\S]*?<\/style>/gi, " ")
    .replace(/<[^>]+>/g, " ");
}

function findRelativeAssets(html) {
  const scanHtml = html.replace(/<script\b[\s\S]*?<\/script>/gi, " ");
  const values = [];
  const attrRegex = /\b(?:src|poster|href)\s*=\s*("[^"]*"|'[^']*'|[^\s>]+)/gi;
  let attrMatch;
  while ((attrMatch = attrRegex.exec(scanHtml))) {
    const value = attrMatch[1].replace(/^['"]|['"]$/g, "");
    if (isRelativeUrl(value)) values.push(value);
  }
  const urlRegex = /url\(\s*(['"]?)([^'")]+)\1\s*\)/gi;
  let urlMatch;
  while ((urlMatch = urlRegex.exec(scanHtml))) {
    if (isRelativeUrl(urlMatch[2])) values.push(urlMatch[2]);
  }
  return [...new Set(values)];
}

function isRelativeUrl(value) {
  return Boolean(value) && !/^(?:data|https?|file|blob|mailto|javascript):/i.test(value) && !value.startsWith("#");
}

async function convertBatchAndDownload(command, htmlPaths, outPath, canvas, args) {
  const stdout = await runDeckops(command, [
    "convert",
    ...htmlPaths,
    "--to",
    "pptx",
    "--width",
    String(canvas.width),
    "--height",
    String(canvas.height),
    "--need-embed-fonts",
    String(args.needEmbedFonts),
    "--timeout",
    args.timeout,
  ]);
  const task = parseJsonOutput(stdout);
  const url = preferPptxUrl(collectUrls(taskResult(task)));
  if (!url) throw new Error(`No PPTX download URL found for DeckFlow convert task ${task.id || "(unknown)"}.`);
  await download(url, outPath);
  return { taskId: task.id, taskType: task.type, outPath };
}

function taskResult(task) {
  return Array.isArray(task) ? task : task?.result;
}

function log(message, jsonMode) {
  if (!jsonMode) console.error(message);
}

function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function escapeHtmlAttr(value) {
  return String(value).replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;");
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  const inputPath = resolve(args.input);
  const outPath = args.out ? resolve(args.out) : null;
  if (!existsSync(inputPath)) throw new Error(`Input file does not exist: ${inputPath}`);
  if (![".html", ".htm"].includes(extname(inputPath).toLowerCase())) {
    throw new Error("Input must be an .html or .htm file.");
  }

  const rawHtml = readFileSync(inputPath, "utf8");
  const tempDir = mkdtempSync(join(tmpdir(), "oh-my-pptx-"));
  const keepTemp = args.keepTemp || args.inspectOnly;
  const globalWarnings = [];

  try {
    const html = prepareHtml(rawHtml, dirname(inputPath), args, globalWarnings);
    const plan = createPages(html, args, inputPath, globalWarnings);
    const htmlPaths = [];

    log(`Prepared ${plan.pages.length} slide HTML file(s) using ${plan.splitMode} split mode.`, args.json);
    log(`Canvas: ${plan.canvas.width}x${plan.canvas.height} (${plan.canvas.source}).`, args.json);

    for (const page of plan.pages) {
      const htmlPath = join(tempDir, `slide-${String(page.index).padStart(3, "0")}.html`);
      writeFileSync(htmlPath, page.html, "utf8");
      htmlPaths.push(htmlPath);
    }

    const inspection = inspectPages(plan.pages, plan.canvas, plan.splitMode, args.scriptMode, globalWarnings);
    const inspectionPath = join(tempDir, "inspection.json");
    writeFileSync(inspectionPath, JSON.stringify(inspection, null, 2), "utf8");

    if (inspection.errors.length > 0 || (args.strict && inspection.warnings.length > 0)) {
      const problems = [...inspection.errors, ...(args.strict ? inspection.warnings : [])];
      throw new Error(`Inspection failed before DeckFlow upload:\n- ${problems.join("\n- ")}\nInspection file: ${inspectionPath}`);
    }

    if (inspection.warnings.length > 0) {
      for (const warning of inspection.warnings) log(`Warning: ${warning}`, args.json);
    }

    let conversion = null;
    if (!args.inspectOnly) {
      const command = deckopsCommand(args.deckops);
      mkdirSync(dirname(outPath), { recursive: true });
      log(`Submitting ${htmlPaths.length} HTML file(s) to DeckFlow batch convert.`, args.json);
      conversion = await convertBatchAndDownload(command, htmlPaths, outPath, plan.canvas, args);
    }

    const summary = {
      ok: true,
      pptx: outPath,
      inspectOnly: args.inspectOnly,
      tempDir: keepTemp ? tempDir : null,
      inspection: keepTemp ? inspectionPath : null,
      slideCount: plan.pages.length,
      splitMode: plan.splitMode,
      scriptMode: args.scriptMode,
      canvas: plan.canvas,
      needEmbedFonts: args.needEmbedFonts,
      warnings: inspection.warnings,
      conversion,
    };
    if (args.json) {
      console.log(JSON.stringify(summary, null, 2));
    } else if (args.inspectOnly) {
      console.log(`Inspection complete: ${inspectionPath}`);
    } else {
      console.log(`PPTX written: ${outPath}`);
    }
  } finally {
    if (keepTemp) {
      log(`Temporary files kept at: ${tempDir}`, args.json);
    } else {
      rmSync(tempDir, { recursive: true, force: true });
    }
  }
}

main().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
