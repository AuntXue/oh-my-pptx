#!/usr/bin/env node

import { spawn, spawnSync } from "node:child_process";
import {
  copyFileSync,
  createWriteStream,
  existsSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  rmSync,
  writeFileSync,
} from "node:fs";
import { tmpdir } from "node:os";
import { dirname, extname, join, resolve } from "node:path";

const DEFAULT_CONVERT_TASK = "convertor.html2pptx";
const DEFAULT_WIDTH = 1920;
const DEFAULT_HEIGHT = 1080;

const USAGE = `
Usage:
  html-to-pptx.mjs <input.html> --out output.pptx [options]

Options:
  --out <file>            Final PPTX path. Required.
  --timeout <sec>         DeckFlow task timeout per task. Default: 300.
  --deckops <cmd>         Deckops command. Default: DECKOPS_BIN, deckops, or npm exec fallback.
  --slide-selector <sel>  Slide selector used for splitting. Default: .slide.
  --width <px>            Override detected slide width.
  --height <px>           Override detected slide height.
  --width-param <name>    DeckFlow width parameter name. Default: width.
  --height-param <name>   DeckFlow height parameter name. Default: height.
  --convert-task <type>   DeckFlow HTML-to-PPTX task type. Default: convertor.html2pptx.
  --param <key=value>     Extra parameter for each HTML-to-PPTX task. Repeatable.
  --keep-temp             Keep temporary split HTML and intermediate PPTX files.
  --json                  Print machine-readable summary. Does not write task JSON files.
  --help                  Show this help.

Environment:
  DECKOPS_BIN, DECKOPS_NPM_SPEC, DECKFLOW_HTML2PPTX_TASK,
  DECKFLOW_WIDTH_PARAM, DECKFLOW_HEIGHT_PARAM
`;

function parseArgs(argv) {
  const args = {
    input: null,
    out: null,
    timeout: "300",
    deckops: process.env.DECKOPS_BIN || null,
    slideSelector: ".slide",
    width: null,
    height: null,
    widthParam: process.env.DECKFLOW_WIDTH_PARAM || "width",
    heightParam: process.env.DECKFLOW_HEIGHT_PARAM || "height",
    convertTask: process.env.DECKFLOW_HTML2PPTX_TASK || DEFAULT_CONVERT_TASK,
    params: [],
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
    } else if (arg === "--slide-selector") {
      args.slideSelector = argv[++i];
    } else if (arg === "--width") {
      args.width = parsePositiveInt(argv[++i], "--width");
    } else if (arg === "--height") {
      args.height = parsePositiveInt(argv[++i], "--height");
    } else if (arg === "--width-param") {
      args.widthParam = argv[++i];
    } else if (arg === "--height-param") {
      args.heightParam = argv[++i];
    } else if (arg === "--convert-task") {
      args.convertTask = argv[++i];
    } else if (arg === "--param") {
      args.params.push(argv[++i]);
    } else if (arg === "--keep-temp") {
      args.keepTemp = true;
    } else if (arg === "--json") {
      args.json = true;
    } else if (!args.input) {
      args.input = arg;
    } else {
      throw new Error(`Unexpected argument: ${arg}`);
    }
  }

  if (!args.input) throw new Error("Missing input HTML file.");
  if (!args.out) throw new Error("Missing required --out output.pptx path.");
  if (!args.timeout || !/^\d+$/.test(args.timeout)) {
    throw new Error("--timeout must be a positive integer.");
  }
  if (!args.slideSelector) throw new Error("--slide-selector cannot be empty.");
  if (!args.widthParam || !args.heightParam) {
    throw new Error("--width-param and --height-param cannot be empty.");
  }
  validateParamList(args.params, "--param");
  return args;
}

function parsePositiveInt(value, name) {
  const number = Number(value);
  if (!Number.isInteger(number) || number <= 0) {
    throw new Error(`${name} must be a positive integer.`);
  }
  return number;
}

function validateParamList(values, flagName) {
  for (const value of values) {
    if (!value || !value.includes("=")) {
      throw new Error(`${flagName} requires key=value.`);
    }
  }
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
    const start = clean.indexOf("{");
    const end = clean.lastIndexOf("}");
    if (start >= 0 && end > start) return JSON.parse(clean.slice(start, end + 1));
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

function htmlShell(originalHtml, slideFragment) {
  const headMatch = originalHtml.match(/<head\b[^>]*>[\s\S]*?<\/head>/i);
  const head = headMatch ? headMatch[0] : "<head><meta charset=\"utf-8\"></head>";
  const bodyAttrMatch = originalHtml.match(/<body\b([^>]*)>/i);
  const bodyAttrs = bodyAttrMatch ? bodyAttrMatch[1] : "";
  return `<!doctype html>\n<html>\n${head}\n<body${bodyAttrs}>\n${slideFragment}\n</body>\n</html>\n`;
}

function splitSlides(html, selector) {
  if (selector === "body") {
    const bodyMatch = html.match(/<body\b[^>]*>([\s\S]*?)<\/body>/i);
    return [bodyMatch ? bodyMatch[1] : html];
  }
  if (!selector.startsWith(".")) {
    throw new Error("Only class selectors such as .slide are supported for --slide-selector.");
  }

  const className = selector.slice(1);
  const fragments = [];
  const tagRegex = /<([a-zA-Z][\w:-]*)(\s[^>]*)?>/g;
  let match;
  while ((match = tagRegex.exec(html))) {
    const [tagText, tagName, attrs = ""] = match;
    if (isVoidOrSelfClosing(tagName, tagText)) continue;
    if (!classListContains(attrs, className)) continue;
    const end = findMatchingClose(html, tagName, tagRegex.lastIndex);
    if (end) fragments.push(html.slice(match.index, end));
  }

  if (fragments.length > 0) return fragments;

  const bodyMatch = html.match(/<body\b[^>]*>([\s\S]*?)<\/body>/i);
  return [bodyMatch ? bodyMatch[1] : html];
}

function classListContains(attrs, className) {
  const classMatch = attrs.match(/\bclass\s*=\s*("[^"]*"|'[^']*'|[^\s>]+)/i);
  if (!classMatch) return false;
  const raw = classMatch[1].replace(/^['"]|['"]$/g, "");
  return raw.split(/\s+/).includes(className);
}

function isVoidOrSelfClosing(tagName, tagText) {
  const voidTags = new Set(["area", "base", "br", "col", "embed", "hr", "img", "input", "link", "meta", "source", "track", "wbr"]);
  return tagText.endsWith("/>") || voidTags.has(tagName.toLowerCase());
}

function findMatchingClose(html, tagName, fromIndex) {
  const escaped = tagName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
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

function detectDimensions(fullHtml, slideFragment, selector, overrides) {
  if (overrides.width && overrides.height) {
    return { width: overrides.width, height: overrides.height, source: "override" };
  }

  const dataSize = detectDataSize(slideFragment);
  const inlineSize = detectInlineStyleSize(slideFragment);
  const cssSize = detectCssRuleSize(fullHtml, selector);
  const viewportSize = detectViewportSize(fullHtml);
  const size = mergeSize(overrides, dataSize) ||
    mergeSize(overrides, inlineSize) ||
    mergeSize(overrides, cssSize) ||
    mergeSize(overrides, viewportSize);

  if (size) return size;
  return { width: overrides.width || DEFAULT_WIDTH, height: overrides.height || DEFAULT_HEIGHT, source: "fallback" };
}

function mergeSize(overrides, detected) {
  if (!detected) return null;
  const width = overrides.width || detected.width;
  const height = overrides.height || detected.height;
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
  const match = tag.match(new RegExp(`\\b${name}\\s*=\\s*("[^"]*"|'[^']*'|[^\\s>]+)`, "i"));
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

function detectCssRuleSize(html, selector) {
  if (!selector.startsWith(".")) return null;
  const className = selector.slice(1);
  const styleBlocks = [...html.matchAll(/<style\b[^>]*>([\s\S]*?)<\/style>/gi)].map((m) => m[1]);
  for (const css of styleBlocks) {
    const ruleRegex = /([^{}]+)\{([^{}]+)\}/g;
    let match;
    while ((match = ruleRegex.exec(css))) {
      const selectors = match[1].split(",").map((item) => item.trim());
      if (!selectors.some((item) => selectorMatchesClass(item, className))) continue;
      const width = declarationPixel(match[2], "width");
      const height = declarationPixel(match[2], "height");
      if (width && height) return { width, height, source: `${selector} CSS rule` };
    }
  }
  return null;
}

function selectorMatchesClass(selector, className) {
  return new RegExp(`(^|[^\\w-])\\.${className.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}($|[^\\w-])`).test(selector);
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

function buildParamArgs(params) {
  const args = [];
  for (const param of params) args.push("--param", param);
  return args;
}

async function runTaskAndDownload(command, taskArgs, outPath) {
  const stdout = await runDeckops(command, taskArgs);
  const task = parseJsonOutput(stdout);
  const url = preferPptxUrl(collectUrls(taskResult(task)));
  if (!url) throw new Error(`No PPTX download URL found for task ${task.id || "(unknown)"}.`);
  await download(url, outPath);
  return { taskId: task.id, taskType: task.type, outPath };
}

async function joinAndDownload(command, pptxPaths, outPath, timeoutSeconds) {
  const stdout = await runDeckops(command, ["join", "--no-wait", ...pptxPaths]);
  const createdTask = parseJsonOutput(stdout);
  if (!createdTask.id) throw new Error("DeckFlow join did not return a task id.");
  const task = await waitForTask(command, createdTask.id, timeoutSeconds);
  const url = preferPptxUrl(collectUrls(taskResult(task)));
  if (!url) throw new Error(`No PPTX download URL found for join task ${task.id || "(unknown)"}.`);
  await download(url, outPath);
  return { taskId: task.id, taskType: task.type, outPath };
}

function taskResult(task) {
  return Array.isArray(task) ? task : task?.result;
}

async function waitForTask(command, taskId, timeoutSeconds) {
  const timeoutMs = Number(timeoutSeconds) * 1000;
  const start = Date.now();
  let lastTask = null;
  while (Date.now() - start <= timeoutMs) {
    const stdout = await runDeckops(command, ["task", "get", taskId]);
    const task = parseJsonOutput(stdout);
    lastTask = task;
    if (task.status === "completed") return task;
    if (task.status === "failed") {
      throw new Error(`DeckFlow task ${taskId} failed: ${task.error || "unknown error"}`);
    }
    await sleep(2000);
  }
  throw new Error(`DeckFlow task ${taskId} did not complete within ${timeoutSeconds}s. Last status: ${lastTask?.status || "unknown"}`);
}

function sleep(ms) {
  return new Promise((resolvePromise) => setTimeout(resolvePromise, ms));
}

function log(message, jsonMode) {
  if (!jsonMode) console.error(message);
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  const inputPath = resolve(args.input);
  const outPath = resolve(args.out);
  if (!existsSync(inputPath)) throw new Error(`Input file does not exist: ${inputPath}`);
  if (![".html", ".htm"].includes(extname(inputPath).toLowerCase())) {
    throw new Error("Input must be an .html or .htm file.");
  }
  if (extname(outPath).toLowerCase() !== ".pptx") throw new Error("--out must end with .pptx.");

  const html = readFileSync(inputPath, "utf8");
  const slideFragments = splitSlides(html, args.slideSelector);
  const tempDir = mkdtempSync(join(tmpdir(), "oh-my-pptx-"));
  const command = deckopsCommand(args.deckops);
  const conversionResults = [];
  const dimensions = [];

  try {
    log(`Detected ${slideFragments.length} slide(s) using ${args.slideSelector}.`, args.json);

    for (let index = 0; index < slideFragments.length; index += 1) {
      const slideNumber = index + 1;
      const dimension = detectDimensions(html, slideFragments[index], args.slideSelector, args);
      dimensions.push(dimension);
      const slideHtml = join(tempDir, `slide-${String(slideNumber).padStart(3, "0")}.html`);
      const slidePptx = join(tempDir, `slide-${String(slideNumber).padStart(3, "0")}.pptx`);
      writeFileSync(slideHtml, htmlShell(html, slideFragments[index]), "utf8");

      log(
        `Converting slide ${slideNumber}/${slideFragments.length} (${dimension.width}x${dimension.height}, ${dimension.source}).`,
        args.json
      );

      const convertParams = [
        `${args.widthParam}=${dimension.width}`,
        `${args.heightParam}=${dimension.height}`,
        ...args.params,
      ];
      const convertTaskArgs = [
        "run",
        args.convertTask,
        slideHtml,
        "--timeout",
        args.timeout,
        ...buildParamArgs(convertParams),
      ];
      conversionResults.push(await runTaskAndDownload(command, convertTaskArgs, slidePptx));
    }

    mkdirSync(dirname(outPath), { recursive: true });
    if (conversionResults.length === 1) {
      copyFileSync(conversionResults[0].outPath, outPath);
    } else {
      log(`Joining ${conversionResults.length} PPTX files in slide order.`, args.json);
      await joinAndDownload(command, conversionResults.map((result) => result.outPath), outPath, args.timeout);
    }

    const summary = {
      ok: true,
      pptx: outPath,
      slideCount: slideFragments.length,
      dimensions: dimensions.map(({ width, height, source }, index) => ({
        slide: index + 1,
        width,
        height,
        source,
      })),
    };
    if (args.json) {
      console.log(JSON.stringify(summary, null, 2));
    } else {
      console.log(`PPTX written: ${outPath}`);
    }
  } finally {
    if (args.keepTemp) {
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
