
// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as commander from "commander";
import { ManifestInfo, readManifestFile } from "office-addin-manifest";
import * as devSettings from "./dev-settings";

export async function clear(manifestPath: string) {
  try {
    const manifest = await readManifestFile(manifestPath);

    validateManifestId(manifest);

    await devSettings.clearDevSettings(manifest.id!);

    console.log("Developer settings have been cleared.");
  } catch (err) {
    logErrorMessage(err);
  }
}

function displaySourceBundleUrl(components: devSettings.SourceBundleUrlComponents) {
  console.log(`host: ${components.host !== undefined ? `"${components.host}"` : '(default: "localhost")'}`);
  console.log(`port: ${components.port !== undefined ? `"${components.port}"` : '(default: "8081")'}`);
  console.log(`path: ${components.path !== undefined ? `"${components.path}"` : "(default)"}`);
  console.log(`extension: ${components.extension !== undefined ? `"${components.extension}"` : '(default: ".bundle")'}`);
  console.log();
  console.log(`Source bundle url: ${components.url}`);
  console.log();
}

export async function disableDebugging(manifestPath: string) {
  try {
    const manifest = await readManifestFile(manifestPath);

    validateManifestId(manifest);

    await devSettings.disableDebugging(manifest.id!);

    console.log("Debugging has been disabled.");
  } catch (err) {
    logErrorMessage(err);
  }
}

export async function disableLiveReload(manifestPath: string) {
  try {
    const manifest = await readManifestFile(manifestPath);

    validateManifestId(manifest);

    await devSettings.disableLiveReload(manifest.id!);

    console.log("Live reload has been disabled.");
  } catch (err) {
    logErrorMessage(err);
  }
}

export async function disableRuntimeLogging() {
  try {
    await devSettings.disableRuntimeLogging();

    console.log("Runtime logging has been disabled.");
  } catch (err) {
    logErrorMessage(err);
  }
}

export async function enableDebugging(manifestPath: string, command: commander.Command) {
  try {
    const manifest = await readManifestFile(manifestPath);

    validateManifestId(manifest);

    await devSettings.enableDebugging(manifest.id!, true, toDebuggingMethod(command.debugMethod));

    console.log("Debugging has been enabled.");
  } catch (err) {
    logErrorMessage(err);
  }
}

export async function enableLiveReload(manifestPath: string) {
  try {
    const manifest = await readManifestFile(manifestPath);

    validateManifestId(manifest);

    await devSettings.enableLiveReload(manifest.id!);

    console.log("Live reload has been enabled.");
  } catch (err) {
    logErrorMessage(err);
  }
}

export async function enableRuntimeLogging(path?: string) {
  try {
    const logPath = await devSettings.enableRuntimeLogging(path);

    console.log(`Runtime logging has been enabled. Log file: ${logPath}`);
  } catch (err) {
    logErrorMessage(err);
  }
}

export async function getSourceBundleUrl(manifestPath: string, command: commander.Command) {
  try {
    const manifest = await readManifestFile(manifestPath);

    validateManifestId(manifest);

    const components: devSettings.SourceBundleUrlComponents = await devSettings.getSourceBundleUrl(manifest.id!);

    displaySourceBundleUrl(components);
  } catch (err) {
    logErrorMessage(err);
  }
}

export async function isDebuggingEnabled(manifestPath: string, command: commander.Command) {
  try {
    const manifest = await readManifestFile(manifestPath);

    validateManifestId(manifest);

    const enabled: boolean = await devSettings.isDebuggingEnabled(manifest.id!);

    console.log(enabled
      ? "Debugging is enabled."
      : "Debugging is not enabled.");
  } catch (err) {
    logErrorMessage(err);
  }
}

export async function isLiveReloadEnabled(manifestPath: string, command: commander.Command) {
  try {
    const manifest = await readManifestFile(manifestPath);

    validateManifestId(manifest);

    const enabled: boolean = await devSettings.isLiveReloadEnabled(manifest.id!);

    console.log(enabled
      ? "Live reload is enabled."
      : "Live reload is not enabled.");
  } catch (err) {
    logErrorMessage(err);
  }
}

export async function isRuntimeLoggingEnabled(command: commander.Command) {
  try {
    const path = await devSettings.isRuntimeLoggingEnabled();

    console.log(path
      ? `Runtime logging is enabled. File: ${path}`
      : "Runtime logging is not enabled.");
  } catch (err) {
    logErrorMessage(err);
  }
}

function logErrorMessage(err: any) {
  console.error(`Error: ${err instanceof Error ? err.message : err}`);
}

export async function setSourceBundleUrl(manifestPath: string, command: commander.Command) {
  try {
    const manifest = await readManifestFile(manifestPath);
    // If the --extension option specifies a value, then command.extension will be a string;
    // otherwise if a value is not specified, command.extension will be the boolean "true".
    // Use the string value (when option value is given); otherwise use 'undefined'.
    // Only --extension allows an empty string as a value; for the others the default is used.
    const host = command.host ? command.host : undefined;
    const port = command.port ? command.port : undefined;
    const path = command.path ? command.path : undefined;
    const extension = (typeof(command.extension) === "string") ? command.extension : undefined;
    const components = new devSettings.SourceBundleUrlComponents(host, port, path, extension);

    validateManifestId(manifest);

    await devSettings.setSourceBundleUrl(manifest.id!, components);

    console.log("Configured source bundle url.");
    displaySourceBundleUrl(components);
  } catch (err) {
    logErrorMessage(err);
  }
}

function toDebuggingMethod(text?: string): devSettings.DebuggingMethod {
  switch (text) {
    case "direct":
      return devSettings.DebuggingMethod.Direct;
    case "web":
      return devSettings.DebuggingMethod.Web;
    case "":
    case null:
    case undefined:
      // preferred debug method
      return devSettings.DebuggingMethod.Web;
    default:
      throw new Error(`Please provide a valid debug method instead of '${text}'.`);
  }
}

function validateManifestId(manifest: ManifestInfo) {
  if (!manifest.id) {
    throw new Error(`The manifest file doesn't contain the id of the Office add-in.`);
  }
}
