
// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as commander from "commander";
import { ManifestInfo, readManifestFile } from "office-addin-manifest";
import {
  addLoopbackExemptionForAppcontainer,
  getAppcontainerName,
  isAppcontainerSupported,
  isLoopbackExemptionForAppcontainer,
  removeLoopbackExemptionForAppcontainer,
} from "./appcontainer";
import * as devSettings from "./dev-settings";

export async function appcontainer(manifestPath: string, command: commander.Command) {
  if (isAppcontainerSupported()) {
    try {
      const manifest = await readManifestFile(manifestPath);
      const sourceLocation = manifest.defaultSettings ? manifest.defaultSettings.sourceLocation : undefined;

      if (sourceLocation === undefined) {
        throw new Error(`The source location could not be retrieved from the manifest.`);
      }

      const name = getAppcontainerName(sourceLocation, false);

      if (command.loopback) {
        console.log(`Appcontainer name: ${name}`);
        try {
          await addLoopbackExemptionForAppcontainer(name);
          const allowed = await isLoopbackExemptionForAppcontainer(name);
          if (allowed) {
            console.log(`Loopback allowed.`);
          } else {
            // if the exemption was not added, the appcontainer name was not found.
            throw new Error("Appcontainer name was not found.");
          }
        } catch (err) {
          throw new Error(`Unable to allow loopback for the appcontainer. \n${err}`);
        }
      } else if (command.preventLoopback) {
        console.log(`Appcontainer name: ${name}`);
        try {
          await removeLoopbackExemptionForAppcontainer(name);
          console.log(`Loopback is no longer allowed.`);
        } catch (err) {
          throw new Error(`Unable to disallow loopback. \n${err}`);
        }
      } else {
        console.log(`Appcontainer name: ${name}`);
        try {
          const allowed = await isLoopbackExemptionForAppcontainer(name);
          console.log(allowed ? "Loopback is allowed." : "Loopback is not allowed.");
        } catch (err) {
          throw new Error(`Unable to determine if appcontainer allows loopback. \n${err}`);
        }
      }
    } catch (err) {
      logErrorMessage(err);
    }
  } else {
    console.log("Appcontainer is not supported.");
  }
}

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

export async function debugging(manifestPath: string, command: commander.Command) {
  if (command.enable) {
    await enableDebugging(manifestPath, command);
  } else if (command.disable) {
    await disableDebugging(manifestPath);
  } else {
    await isDebuggingEnabled(manifestPath);
  }
}

function displaySourceBundleUrl(components: devSettings.SourceBundleUrlComponents) {
  console.log(`host: ${components.host !== undefined ? `"${components.host}"` : '"localhost" (default)'}`);
  console.log(`port: ${components.port !== undefined ? `"${components.port}"` : '"8081" (default)'}`);
  console.log(`path: ${components.path !== undefined ? `"${components.path}"` : "(default)"}`);
  console.log(`extension: ${components.extension !== undefined ? `"${components.extension}"` : '".bundle" (default)'}`);
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

    console.log(`Runtime logging has been enabled. File: ${logPath}`);
  } catch (err) {
    logErrorMessage(err);
  }
}

export async function getSourceBundleUrl(manifestPath: string) {
  try {
    const manifest = await readManifestFile(manifestPath);

    validateManifestId(manifest);

    const components: devSettings.SourceBundleUrlComponents = await devSettings.getSourceBundleUrl(manifest.id!);

    displaySourceBundleUrl(components);
  } catch (err) {
    logErrorMessage(err);
  }
}

export async function isDebuggingEnabled(manifestPath: string) {
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

export async function isLiveReloadEnabled(manifestPath: string) {
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

export async function isRuntimeLoggingEnabled() {
  try {
    const path = await devSettings.getRuntimeLoggingPath();

    console.log(path
      ? `Runtime logging is enabled. File: ${path}`
      : "Runtime logging is not enabled.");
  } catch (err) {
    logErrorMessage(err);
  }
}

export async function liveReload(manifestPath: string, command: commander.Command) {
  if (command.enable) {
    await enableLiveReload(manifestPath);
  } else if (command.disable) {
    await disableLiveReload(manifestPath);
  } else {
    await isLiveReloadEnabled(manifestPath);
  }
}

function logErrorMessage(err: any) {
  console.error(`Error: ${err instanceof Error ? err.message : err}`);
}

function parseStringCommandOption(optionValue: any): string | undefined {
  return (typeof(optionValue) === "string") ? optionValue : undefined;
}

export async function runtimeLogging(command: commander.Command) {
  if (command.enable) {
    const path: string | undefined = (typeof(command.enable) === "string") ? command.enable : undefined;
    await enableRuntimeLogging(path);
  } else if (command.disable) {
    await disableRuntimeLogging();
  } else {
    await isRuntimeLoggingEnabled();
  }
}

export async function setSourceBundleUrl(manifestPath: string, command: commander.Command) {
  try {
    const manifest = await readManifestFile(manifestPath);
    const host = parseStringCommandOption(command.host);
    const port = parseStringCommandOption(command.port);
    const path = parseStringCommandOption(command.path);
    const extension = parseStringCommandOption(command.extension);
    const components = new devSettings.SourceBundleUrlComponents(host, port, path, extension);

    validateManifestId(manifest);

    await devSettings.setSourceBundleUrl(manifest.id!, components);

    console.log("Configured source bundle url.");
    displaySourceBundleUrl(await devSettings.getSourceBundleUrl(manifest.id!));
  } catch (err) {
    logErrorMessage(err);
  }
}

export async function sourceBundleUrl(manifestPath: string, command: commander.Command) {
  if (command.host !== undefined || command.port !== undefined || command.path !== undefined || command.extension !== undefined) {
    await setSourceBundleUrl(manifestPath, command);
  } else {
    await getSourceBundleUrl(manifestPath);
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
