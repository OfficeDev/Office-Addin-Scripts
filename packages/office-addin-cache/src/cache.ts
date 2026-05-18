// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import fs from "fs/promises";
import path from "path";
import os from "os";
import { execFile } from "child_process";
import { promisify } from "util";
import readline from "readline";

/* global process console setTimeout setInterval clearInterval */

const execFileAsync = promisify(execFile);

/* -------------------- Types -------------------- */

type Platform = string;

type RunningProcess = { kind: "process"; name: string } | { kind: "app"; name: string };

type CacheTarget = {
  label: string;
  dir: string;
};

export type Options = {
  verbose: boolean;
  forceClose: boolean;
};

type InlineSpinner = {
  enabled: boolean;
  start: () => void;
  stop: () => void;
};

/* -------------------- Utils -------------------- */

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

/* -------------------- Spinner -------------------- */

function createInlineSpinner(): InlineSpinner {
  const enabled: boolean = process.stdout.isTTY ?? false;
  const frames: string[] = ["|", "/", "-", "|", "\\", "-"];
  const intervalMs = 80;

  let timer: ReturnType<typeof setInterval> | null = null;
  let idx = 0;
  let active = false;
  let showing = false;

  function hideCursor(): void {
    process.stdout.write("\x1B[?25l");
  }

  function showCursor(): void {
    process.stdout.write("\x1B[?25h");
  }

  function tick(): void {
    if (!active) return;
    const ch = frames[idx];
    idx = (idx + 1) % frames.length;
    if (showing) process.stdout.write("\b");
    process.stdout.write(ch);
    showing = true;
  }

  function start(): void {
    if (!enabled || active) return;
    active = true;
    showing = false;
    idx = 0;
    hideCursor();
    timer = setInterval(tick, intervalMs);
  }

  function stop(): void {
    if (!enabled || !active) return;
    if (timer) clearInterval(timer);
    timer = null;
    if (showing) process.stdout.write("\b \b");
    showing = false;
    active = false;
    showCursor();
  }

  return { enabled, start, stop };
}

/* -------------------- Office Application Processes -------------------- */

const WIN_PROCS: string[] = [
  "WINWORD.EXE",
  "EXCEL.EXE",
  "POWERPNT.EXE",
  // Outlook classic
  "OUTLOOK.EXE",
  "ONENOTE.EXE",
  // New Outlook
  "OLK.EXE",
  "WINPROJ.EXE",
];

const MAC_APPS: string[] = [
  "Microsoft Word",
  "Microsoft Excel",
  "Microsoft PowerPoint",
  "Microsoft Outlook",
  "Microsoft OneNote",
  "Microsoft Project",
];

function allOfficeCandidatesForPlatform(): string[] {
  return process.platform === "win32" ? WIN_PROCS : MAC_APPS;
}

async function detectOfficeRunning(candidates?: string[]): Promise<RunningProcess[]> {
  const platform: Platform = process.platform;
  const officeAppProcNames: string[] =
    candidates && candidates.length ? candidates : allOfficeCandidatesForPlatform();

  if (platform === "win32") {
    const { stdout } = await execFileAsync("tasklist", ["/FO", "CSV", "/NH"]);
    const lines = stdout.split("\n");
    const wanted = new Set<string>(officeAppProcNames.map((n) => n.toUpperCase()));
    const running: RunningProcess[] = [];

    for (const line of lines) {
      if (!line.trim()) continue;
      const cols = line.split('","').map((s: string) => s.replace(/^"|"$/g, ""));
      const imageName = cols[0]?.toUpperCase();
      if (imageName && wanted.has(imageName)) {
        running.push({ kind: "process", name: imageName });
      }
    }

    return running;
  }

  if (platform === "darwin") {
    const { stdout } = await execFileAsync("ps", ["-ax"]);
    const lines = stdout.split("\n");
    const running: RunningProcess[] = [];
    const wanted = new Set<string>(officeAppProcNames);

    for (const line of lines) {
      if (!line.trim()) continue;
      const parts = line.trim().split(/\s+/);
      const command = parts.slice(4).join(" ");
      const base = path.basename(command);
      if (wanted.has(base)) {
        running.push({ kind: "app", name: base });
      }
    }

    return running;
  }

  return [];
}

async function forceCloseOfficeApps(
  runningList: RunningProcess[],
  verbose: boolean
): Promise<void> {
  const platform: Platform = process.platform;

  if (platform === "win32") {
    for (const proc of runningList) {
      try {
        if (verbose) console.log(`Force-closing: ${proc.name}`);
        await execFileAsync("taskkill", ["/F", "/IM", proc.name, "/T"]);
      } catch {
        if (verbose) console.log(`Failed to force-close: ${proc.name}`);
      }
    }
    return;
  }

  if (platform === "darwin") {
    for (const proc of runningList) {
      try {
        if (verbose) console.log(`Force-killing: ${proc.name}`);
        await execFileAsync("pkill", ["-9", "-f", proc.name]);
      } catch {
        if (verbose) console.log(`Failed to force-kill: ${proc.name}`);
      }
    }
  }
}

function formatRunningList(running: RunningProcess[]): string[] {
  const processDisplayNames = new Map<string, string>([
    ["WINWORD.EXE", "Word"],
    ["EXCEL.EXE", "Excel"],
    ["POWERPNT.EXE", "PowerPoint"],
    ["OUTLOOK.EXE", "Outlook"],
    ["ONENOTE.EXE", "OneNote"],
    ["OLK.EXE", "New Outlook"],
    ["WINPROJ.EXE", "Project"],
  ]);

  return running.map((p) => processDisplayNames.get(p.name) ?? p.name);
}

async function scanWithSpinner(
  message: string,
  spinner: InlineSpinner,
  candidates?: string[]
): Promise<RunningProcess[]> {
  process.stdout.write(`${message} `);
  spinner.start();
  try {
    return await detectOfficeRunning(candidates);
  } finally {
    spinner.stop();
    process.stdout.write("\n");
  }
}

async function verifyClosed(previousNames: string[], spinner: InlineSpinner): Promise<boolean> {
  process.stdout.write("Verifying that Office applications have been closed ... ");
  spinner.start();

  const start = Date.now();
  const timeoutMs = 20000;

  try {
    while (true) {
      const still = await detectOfficeRunning(previousNames);
      if (still.length === 0) return true;
      if (Date.now() - start > timeoutMs) return false;
      await sleep(250);
    }
  } finally {
    spinner.stop();
    process.stdout.write("\n");
  }
}

async function promptYesNo(question: string): Promise<boolean> {
  if (!process.stdin.isTTY) return false;

  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
  });

  const answer: string = await new Promise((resolve) => rl.question(question, resolve));
  rl.close();
  return /^y(es)?$/i.test(answer.trim());
}

/* -------------------- Cache -------------------- */

export async function getCacheTargets(): Promise<CacheTarget[]> {
  const platform: Platform = process.platform;
  const targets: CacheTarget[] = [];

  if (platform === "win32") {
    const base: string = process.env.USERPROFILE ?? "";

    targets.push({
      label: "WEF",
      dir: path.join(process.env.LOCALAPPDATA ?? "", "Microsoft", "Office", "16.0", "Wef"),
    });

    targets.push({
      label: "WebView Cache",
      dir: path.join(
        base,
        "AppData",
        "Local",
        "Packages",
        "Microsoft.Win32WebViewHost_cw5n1h2txyewy"
      ),
    });

    targets.push({
      label: "Outlook Hub App Cache",
      dir: path.join(base, "AppData", "Local", "Microsoft", "Outlook", "HubAppFileCache"),
    });

    return targets;
  }

  if (platform === "darwin") {
    const home = os.homedir();
    return [
      {
        label: "OsfWebHost",
        dir: path.join(home, "Library/Containers/com.Microsoft.OsfWebHost/Data"),
      },
    ];
  }

  return targets;
}

async function clearTargets(targets: CacheTarget[], verbose: boolean): Promise<void> {
  for (const target of targets) {
    if (verbose) console.log(`Clearing: ${target.label} (${target.dir})`);
    try {
      await fs.rm(target.dir, { recursive: true, force: true });
      console.log(`Cleared: ${target.label}`);
    } catch {
      console.log(`Skipped: ${target.label}`);
      console.log(`Please clear this folder manually: ${target.dir}`);
    }
  }
}

/* -------------------- Primary Export -------------------- */

export async function clearCache(options: Options): Promise<void> {
  const spinner: InlineSpinner = createInlineSpinner();

  const running: RunningProcess[] = await scanWithSpinner(
    "Looking for running Office applications ...",
    spinner,
    allOfficeCandidatesForPlatform()
  );

  if (running.length > 0) {
    const friendlyAppNames: string[] = formatRunningList(running);

    console.log("Office applications are currently running:");
    friendlyAppNames.forEach((name) => console.log(`  - ${name}`));

    if (options.forceClose) {
      console.log("\nForce-close option specified. Closing Office applications automatically...");
    } else {
      console.log(
        "\nIf you choose to force-close them, the tool will continue and clear the cache automatically."
      );
      console.log(
        "If you choose not to, close all Office applications manually before rerunning this tool."
      );
    }

    const doForce = options.forceClose || (await promptYesNo("\nForce-close? (y/N): "));

    if (!doForce) {
      console.log("\nNo changes made.");
      console.log("Please close all Office apps and rerun the command.");
      process.exitCode = 1;
      return;
    }

    await forceCloseOfficeApps(running, options.verbose);

    const closed = await verifyClosed(
      running.map((process) => process.name),
      spinner
    );

    if (!closed) {
      console.log("\nUnable to close all Office applications.");
      console.log("\nPlease close them manually and then rerun this tool.");
      process.exitCode = 1;
      return;
    }

    console.log("All previously detected Office apps are no longer running.");
    console.log("Proceeding to clear the Office Add-ins cache...");
  }

  const targets: CacheTarget[] = await getCacheTargets();
  await clearTargets(targets, options.verbose);

  console.log("\nDone.");
}
