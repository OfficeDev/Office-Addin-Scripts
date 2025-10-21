// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import childProcess from "child_process";
import inquirer from "inquirer";
import { OfficeAddinManifest } from "office-addin-manifest";
import { URL } from "whatwg-url";
import { ExpectedError } from "office-addin-usage-data";

/* global process */

export const EdgeBrowserAppcontainerName: string = "Microsoft.MicrosoftEdge_8wekyb3d8bbwe";
export const EdgeWebViewAppcontainerName: string = "Microsoft.win32webviewhost_cw5n1h2txyewy";
export const EdgeBrowserName: string = "Microsoft Edge Web Browser";
export const EdgeWebViewName: string = "Microsoft Edge WebView";

/**
 * Adds a loopback exemption for the appcontainer.
 * @param name Appcontainer name
 * @throws Error if platform does not support appcontainers.
 */
export function addLoopbackExemptionForAppcontainer(name: string): Promise<void> {
  if (!isAppcontainerSupported()) {
    throw new ExpectedError(`Platform not supported: ${process.platform}.`);
  }

  return new Promise((resolve, reject) => {
    const command = `CheckNetIsolation.exe LoopbackExempt -a -n=${name}`;

    childProcess.exec(command, (error, stdout) => {
      if (error) {
        reject(stdout);
      } else {
        resolve();
      }
    });
  });
}

/**
 * Returns whether appcontainer is supported on the current platform.
 * @returns True if platform supports using appcontainer; false otherwise.
 */
export function isAppcontainerSupported() {
  return process.platform === "win32";
}

/**
 * Adds a loopback exemption for the appcontainer.
 * @param name Appcontainer name
 * @throws Error if platform does not support appcontainers.
 */
export function isLoopbackExemptionForAppcontainer(name: string): Promise<boolean> {
  if (!isAppcontainerSupported()) {
    throw new ExpectedError(`Platform not supported: ${process.platform}.`);
  }

  return new Promise((resolve, reject) => {
    const command = `CheckNetIsolation.exe LoopbackExempt -s`;

    childProcess.exec(command, (error, stdout) => {
      if (error) {
        reject(stdout);
      } else {
        const expr = new RegExp(`Name: ${name}`, "i");
        const found: boolean = expr.test(stdout);
        resolve(found);
      }
    });
  });
}

/**
 * Returns the name of the appcontainer used to run an Office Add-in.
 * @param sourceLocation Source location of the Office Add-in.
 * @param isFromStore True if installed from the Store; false otherwise.
 */
export function getAppcontainerName(sourceLocation: string, isFromStore = false): string {
  const url: URL = new URL(sourceLocation);
  const origin: string = url.origin;
  const addinType: number = isFromStore ? 0 : 1; // 0 if from Office Add-in store, 1 otherwise.
  const guid: string = "04ACA5EC-D79A-43EA-AB47-E50E47DD96FC";

  // Appcontainer name format is "{addinType}_{origin}{guid}".
  // Replace characters ":" and "/" with "_".
  const name = `${addinType}_${origin}${guid}`.replace(/[://]/g, "_");

  return name;
}

export async function getUserConfirmation(name: string): Promise<boolean> {
  const answers = await inquirer.prompt({
    message: `Allow localhost loopback for ${name}?`,
    name: "didUserConfirm",
    type: "confirm",
  });
  return (answers as any).didUserConfirm;
}

export async function getAppcontainerNameFromManifestPath(manifestPath: string): Promise<string> {
  switch (manifestPath.toLowerCase()) {
    case "edgewebview":
      return EdgeWebViewAppcontainerName;
    case "edgewebbrowser":
    case "edge":
      return EdgeBrowserAppcontainerName;
    default:
      return await getAppcontainerNameFromManifest(manifestPath);
  }
}

export function getDisplayNameFromManifestPath(manifestPath: string): string {
  switch (manifestPath.toLowerCase()) {
    case "edgewebview":
      return EdgeWebViewName;
    case "edgewebbrowser":
    case "edge":
      return EdgeBrowserName;
    default:
      return manifestPath;
  }
}

export async function ensureLoopbackIsEnabled(
  manifestPath: string,
  askForConfirmation: boolean = true
): Promise<boolean> {
  const name = await getAppcontainerNameFromManifestPath(manifestPath);
  let isEnabled = await isLoopbackExemptionForAppcontainer(name);

  if (!isEnabled) {
    if (
      !askForConfirmation ||
      (await getUserConfirmation(getDisplayNameFromManifestPath(manifestPath)))
    ) {
      await addLoopbackExemptionForAppcontainer(name);
      isEnabled = true;
    }
  }

  return isEnabled;
}

/**
 * Returns the name of the appcontainer used to run an Office Add-in.
 * @param manifestPath Path of the manifest file.
 */
export async function getAppcontainerNameFromManifest(manifestPath: string): Promise<string> {
  const manifest = await OfficeAddinManifest.readManifestFile(manifestPath);
  const sourceLocation = manifest.defaultSettings
    ? manifest.defaultSettings.sourceLocation
    : undefined;

  if (sourceLocation === undefined) {
    throw new ExpectedError(`The source location could not be retrieved from the manifest.`);
  }

  return getAppcontainerName(sourceLocation, false);
}

/**
 * Removes a loopback exemption for the appcontainer.
 * @param name Appcontainer name
 * @throws Error if platform doesn't support appcontainers.
 */
export function removeLoopbackExemptionForAppcontainer(name: string): Promise<void> {
  if (!isAppcontainerSupported()) {
    throw new ExpectedError(`Platform not supported: ${process.platform}.`);
  }

  return new Promise((resolve, reject) => {
    const command = `CheckNetIsolation.exe LoopbackExempt -d -n=${name}`;

    childProcess.exec(command, (error, stdout) => {
      if (error) {
        reject(stdout);
      } else {
        resolve();
      }
    });
  });
}
