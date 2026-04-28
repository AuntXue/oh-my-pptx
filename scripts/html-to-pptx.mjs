#!/usr/bin/env node

import { spawn, spawnSync } from "node:child_process";
import { existsSync, mkdirSync, writeFileSync, createWriteStream } from "node:fs";
import { dirname, extname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const USAGE = `
Usage:
  html-to-pptx.mjs <input.html> [--out output.pptx] [--timeout seconds]

Options:
  --out <file>       Local PPTX path to write when deckops returns a download URL.
  --timeout <sec>    Deckops task timeout. Default: 300.
  --deckops <cmd>    Deckops command to run. Default: DECKOPS_BIN, deckops, or npm exec fallback.
                    Set DECKOPS_NPM_SPEC to change fallback package version.
  --no-download      Do not download result URLs; only save task JSON.
  --help             Show this help.
`;

function parseArgs(argv) {
  const args = {
    input: null,
    out: null,
    timeout: "300",
    deckops: process.env.DECKOPS_BIN || null,
    download: true,
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
    } else if (arg === "--no-download") {
      args.download = false;
    } else if (!args.input) {
      args.input = arg;
    } else {
      throw new Error(`Unexpected argument: ${arg}`);
    }
  }

  if (!args.input) throw new Error("Missing input HTML file.");
  if (!args.timeout || !/^\d+$/.test(args.timeout)) {
    throw new Error("--timeout must be a positive integer.");
  }
  return args;
}

function commandExists(cmd) {
  const result = spawnSync(cmd, ["--help"], { stdio: "ignore" });
  return result.status === 0 || result.status === 1;
}

function deckopsCommand(explicit) {
  if (explicit) return { cmd: explicit, baseArgs: [] };
  if (commandExists("deckops")) return { cmd: "deckops", baseArgs: [] };
  return { cmd: "npm", baseArgs: ["exec", "-y", process.env.DECKOPS_NPM_SPEC || "deckops@0.2.1", "--"] };
}

function runDeckops({ cmd, baseArgs }, input, timeout) {
  const args = [
    ...baseArgs,
    "--json",
    "convert",
    input,
    "--to",
    "pptx",
    "--timeout",
    timeout,
  ];

  return new Promise((resolvePromise, reject) => {
    const child = spawn(cmd, args, { stdio: ["ignore", "pipe", "pipe"] });
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
  try {
    return JSON.parse(stdout);
  } catch {
    const start = stdout.indexOf("{");
    const end = stdout.lastIndexOf("}");
    if (start >= 0 && end > start) {
      return JSON.parse(stdout.slice(start, end + 1));
    }
    throw new Error(`Could not parse deckops JSON output:\n${stdout}`);
  }
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
  if (!response.ok) {
    throw new Error(`Download failed: ${response.status} ${response.statusText}`);
  }
  mkdirSync(dirname(outPath), { recursive: true });
  const file = createWriteStream(outPath);
  await new Promise((resolvePromise, reject) => {
    response.body.pipeTo(
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
    ).catch(reject);
  });
}

function defaultOutPath(input) {
  return input.replace(/\.html?$/i, ".pptx");
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  const inputPath = resolve(args.input);
  if (!existsSync(inputPath)) throw new Error(`Input file does not exist: ${inputPath}`);
  if (![".html", ".htm"].includes(extname(inputPath).toLowerCase())) {
    throw new Error("Input must be an .html or .htm file.");
  }

  const outPath = resolve(args.out || defaultOutPath(inputPath));
  if (extname(outPath).toLowerCase() !== ".pptx") {
    throw new Error("--out must end with .pptx.");
  }
  const taskJsonPath = outPath.replace(/\.pptx$/i, ".task.json");
  const command = deckopsCommand(args.deckops);

  console.error(`Running: ${command.cmd} ${[...command.baseArgs, "--json", "convert", inputPath, "--to", "pptx"].join(" ")}`);
  const stdout = await runDeckops(command, inputPath, args.timeout);
  const task = parseJsonOutput(stdout);
  writeFileSync(taskJsonPath, `${JSON.stringify(task, null, 2)}\n`, "utf8");

  const urls = collectUrls(task.result);
  const url = preferPptxUrl(urls);
  if (args.download && url) {
    await download(url, outPath);
    console.log(JSON.stringify({ ok: true, pptx: outPath, taskJson: taskJsonPath, taskId: task.id }, null, 2));
  } else {
    console.log(JSON.stringify({
      ok: Boolean(task && task.status === "completed"),
      pptx: null,
      taskJson: taskJsonPath,
      taskId: task.id,
      note: url ? "Download skipped by --no-download." : "No downloadable URL found in task.result.",
    }, null, 2));
    if (!url && args.download) process.exitCode = 2;
  }
}

main().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
