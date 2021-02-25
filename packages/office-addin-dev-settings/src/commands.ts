
// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as commander from "commander";
import { logErrorMessage, parseNumber } from "office-addin-cli";
import { ManifestInfo, OfficeApp, parseOfficeApp, readManifestFile } from "office-addin-manifest";
import {
  ensureLoopbackIsEnabled,
  getAppcontainerNameFromManifestPath,
  isAppcontainerSupported,
  isLoopbackExemptionForAppcontainer,
  removeLoopbackExemptionForAppcontainer,
} from "./appcontainer";
import { AppType, parseAppType } from "./appType";
import * as devSettings from "./dev-settings";
import { sideloadAddIn } from "./sideload";
import { usageDataObject } from './defaults';
import { ExpectedError } from "office-addin-usage-data";

export async function appcontainer(manifestPath: string, command: commander.Command) {
  if (isAppcontainerSupported()) {
    try {
      if (command.loopback) {
        try {
          const askForConfirmation: boolean = (command.yes !== true);
          const allowed = await ensureLoopbackIsEnabled(manifestPath, askForConfirmation);
          console.log(allowed ? "Loopback is allowed." : "Loopback is not allowed.");
        } catch (err) {
          throw new Error(`Unable to allow loopback for the appcontainer. \n${err}`);
        }
      } else if (command.preventLoopback) {
        try {
          const name = await getAppcontainerNameFromManifestPath(manifestPath);
          await removeLoopbackExemptionForAppcontainer(name);
          console.log(`Loopback is no longer allowed.`);
        } catch (err) {
          throw new Error(`Unable to disallow loopback. \n${err}`);
        }
      } else {
        try {
          const name = await getAppcontainerNameFromManifestPath(manifestPath);
          const allowed = await isLoopbackExemptionForAppcontainer(name);
          console.log(allowed ? "Loopback is allowed." : "Loopback is not allowed.");
        } catch (err) {
          throw new Error(`Unable to determine if appcontainer allows loopback. \n${err}`);
        }
      }
      usageDataObject.reportSuccess("appcontainer");
    } catch (err) {
      usageDataObject.reportException("appcontainer", err);
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
    usageDataObject.reportSuccess("clear");
  } catch (err) {
    usageDataObject.reportException("clear", err);
    logErrorMessage(err);
  }
}

export async function debugging(manifestPath: string, command: commander.Command) {
  try {
    if (command.enable) {
      await enableDebugging(manifestPath, command);
    } else if (command.disable) {
      await disableDebugging(manifestPath);
    } else {
      await isDebuggingEnabled(manifestPath);
    }
    usageDataObject.reportSuccess("debugging");
  } catch (err) {
    usageDataObject.reportException("debugging", err);
    logErrorMessage(err);
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
    usageDataObject.reportSuccess("disableDebugging()");
  } catch (err) {
    usageDataObject.reportException("disableDebugging()", err);
    logErrorMessage(err);
  }
}    

export async function disableLiveReload(manifestPath: string) {
  try {
    const manifest = await readManifestFile(manifestPath);

    validateManifestId(manifest);

    await devSettings.disableLiveReload(manifest.id!);

    console.log("Live reload has been disabled.");
    usageDataObject.reportSuccess("disableLiveReload()");
  } catch (err) {
    usageDataObject.reportException("disableLiveReload()", err);
    logErrorMessage(err);
  }
}

export async function disableRuntimeLogging() {
  try {
    await devSettings.disableRuntimeLogging();

    console.log("Runtime logging has been disabled.");
    usageDataObject.reportSuccess("disableRuntimeLogging()");
  } catch (err) {
    usageDataObject.reportException("disableRuntimeLogging()", err);
    logErrorMessage(err);
  }
}

export async function enableDebugging(manifestPath: string, command: commander.Command) {
  try {
    const manifest = await readManifestFile(manifestPath);

    validateManifestId(manifest);

    await devSettings.enableDebugging(manifest.id!, true, toDebuggingMethod(command.debugMethod), command.openDevTools);

    console.log("Debugging has been enabled.");
    usageDataObject.reportSuccess("enableDebugging()");
  } catch (err) {
    usageDataObject.reportException("enableDebugging()", err);
    logErrorMessage(err);
  }
}

export async function enableLiveReload(manifestPath: string) {
  try {
    const manifest = await readManifestFile(manifestPath);

    validateManifestId(manifest);

    await devSettings.enableLiveReload(manifest.id!);

    console.log("Live reload has been enabled.");
    usageDataObject.reportSuccess("enableLiveReload()");
  } catch (err) {
    usageDataObject.reportException("enableLiveReload()", err);
    logErrorMessage(err);
  }
}

export async function enableRuntimeLogging(path?: string) {
  try {
    const logPath = await devSettings.enableRuntimeLogging(path);

    console.log(`Runtime logging has been enabled. File: ${logPath}`);
    usageDataObject.reportSuccess("enableRuntimeLogging()");
  } catch (err) {
    usageDataObject.reportException("enableRuntimeLogging()", err);
    logErrorMessage(err);
  }
}

export async function getSourceBundleUrl(manifestPath: string) {
  try {
    const manifest = await readManifestFile(manifestPath);

    validateManifestId(manifest);

    const components: devSettings.SourceBundleUrlComponents = await devSettings.getSourceBundleUrl(manifest.id!);

    displaySourceBundleUrl(components);
    usageDataObject.reportSuccess("getSourceBundleUrl()");
  } catch (err) {
    usageDataObject.reportException("getSourceBundleUrl()", err);
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
      usageDataObject.reportSuccess("isDebuggingEnabled()");
  } catch (err) {
    usageDataObject.reportException("isDebuggingEnabled()", err);
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
    usageDataObject.reportSuccess("isLiveReloadEnabled()");
  } catch (err) {
    usageDataObject.reportException("isLiveReloadEnabled()", err);
    logErrorMessage(err);
  }
}
  
export async function isRuntimeLoggingEnabled() {
  try {
    const path = await devSettings.getRuntimeLoggingPath();

    console.log(path
      ? `Runtime logging is enabled. File: ${path}`
      : "Runtime logging is not enabled.");
    usageDataObject.reportSuccess("isRuntimeLoggingEnabled()");
  } catch (err) {
    usageDataObject.reportException("isRuntimeLoggingEnabled()", err);
    logErrorMessage(err);
  }
}

export async function liveReload(manifestPath: string, command: commander.Command) {
  try{
    if (command.enable) {
      await enableLiveReload(manifestPath);
    } else if (command.disable) {
      await disableLiveReload(manifestPath);
    } else {
      await isLiveReloadEnabled(manifestPath);
    }
    usageDataObject.reportSuccess("liveReload");
  } catch (err) {
    usageDataObject.reportException("liveReload", err);
    logErrorMessage(err);
  }
}

function parseStringCommandOption(optionValue: any): string | undefined {
  return (typeof(optionValue) === "string") ? optionValue : undefined;
}

export function parseWebViewType(webViewString?: string): devSettings.WebViewType | undefined {
  switch (webViewString ? webViewString.toLowerCase() : undefined) {
    case "ie":
    case "ie11":
    case "internet explorer":
    case "internetexplorer":
      return devSettings.WebViewType.IE;
    case "edgelegacy":
    case "edge-legacy":
    case "spartan":
      return devSettings.WebViewType.Edge;
    case "edge chromium":
    case "edgechromium":
    case "edge-chromium":
    case "anaheim":
    case "edge":
      return devSettings.WebViewType.EdgeChromium;
    case "default":
    case "":
    case null:
    case undefined:
      return undefined;
    default:
      throw new ExpectedError(`Please select a valid web view type instead of '${webViewString!}'.`);
  }
}

export async function register(manifestPath: string, command: commander.Command) {
  try {
    await devSettings.registerAddIn(manifestPath);
    usageDataObject.reportSuccess("register");
  } catch (err) {
    usageDataObject.reportException("register", err);
    logErrorMessage(err);
  }
}

export async function registered(command: commander.Command) {
  try {
    const registeredAddins = await devSettings.getRegisterAddIns();

    if (registeredAddins.length > 0) {
      for (const addin of registeredAddins) {
        let id: string = addin.id;

        if (!id && addin.manifestPath) {
          try {
            const manifest = await readManifestFile(addin.manifestPath);
            id = manifest.id || "";
          } catch (err) {
            // ignore errors
          }
        }

        console.log(`${id ? id + " " : ""}${addin.manifestPath}`);
      }
    } else {
      console.log("No add-ins are registered.");
    }
    usageDataObject.reportSuccess("registered");
  } catch (err) {
    usageDataObject.reportException("registered", err);
    logErrorMessage(err);
  }
}

export async function runtimeLogging(command: commander.Command) {
  try {
    if (command.enable) {
      const path: string | undefined = (typeof(command.enable) === "string") ? command.enable : undefined;
      await enableRuntimeLogging(path);
    } else if (command.disable) {
      await disableRuntimeLogging();
    } else {
      await isRuntimeLoggingEnabled();
    }
    usageDataObject.reportSuccess("runtimeLogging");
  } catch (err) {
    usageDataObject.reportException("runtimeLogging", err);
    logErrorMessage(err);
  }
}

export async function sideload(manifestPath: string, type: string | undefined, command: commander.Command) {
  try {
    const app: OfficeApp | undefined = command.app ? parseOfficeApp(command.app) : undefined;
    const canPrompt = true;
    const document: string | undefined = command.document ? command.document : undefined;
    const appType: AppType | undefined = parseAppType(type || process.env.npm_package_config_app_platform_to_debug);

    await sideloadAddIn(manifestPath, app, canPrompt, appType, document);
    usageDataObject.reportSuccess("sideload");
  } catch (err) {
    usageDataObject.reportException("sideload", err);
    logErrorMessage(err);
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
    usageDataObject.reportSuccess("setSourceBundleUrl()");
  } catch (err) {
    usageDataObject.reportException("setSourceBundleUrl()", err);
    logErrorMessage(err);
  }
}

export async function sourceBundleUrl(manifestPath: string, command: commander.Command) {
  try {
    if (command.host !== undefined || command.port !== undefined || command.path !== undefined || command.extension !== undefined) {
      await setSourceBundleUrl(manifestPath, command);
    } else {
      await getSourceBundleUrl(manifestPath);
    }
    usageDataObject.reportSuccess("sourceBundleUrl");
  } catch (err) {
    usageDataObject.reportException("sourceBundleUrl", err);
    logErrorMessage(err);
  }
}

function parseDevServerPort(optionValue: any): number | undefined {
  const devServerPort = parseNumber(optionValue, "--dev-server-port should specify a number.");

  if (devServerPort !== undefined) {
      if (!Number.isInteger(devServerPort)) {
          throw new ExpectedError("--dev-server-port should be an integer.");
      }
      if ((devServerPort < 0) || (devServerPort > 65535)) {
          throw new ExpectedError("--dev-server-port should be between 0 and 65535.");
      }
  }

  return devServerPort;
}

function toDebuggingMethod(text?: string): devSettings.DebuggingMethod {
  switch (text) {
    case "direct":
      return devSettings.DebuggingMethod.Direct;
    case "proxy":
      return devSettings.DebuggingMethod.Proxy;
    case "":
    case null:
    case undefined:
      // preferred debug method
      return devSettings.DebuggingMethod.Direct;
    default:
      throw new ExpectedError(`Please provide a valid debug method instead of '${text}'.`);
  }
}

export async function unregister(manifestPath: string, command: commander.Command) {
  try {
    if (manifestPath === "all") {
      await devSettings.unregisterAllAddIns();
    } else {
      await devSettings.unregisterAddIn(manifestPath);
    }
    usageDataObject.reportSuccess("unregister");
  } catch (err) {
    usageDataObject.reportException("unregister", err);
    logErrorMessage(err);
  }
}

function validateManifestId(manifest: ManifestInfo) {
  if (!manifest.id) {
    throw new ExpectedError(`The manifest file doesn't contain the id of the Office Add-in.`);
  }
}

export async function webView(manifestPath: string, webViewString?: string) {
  try {
    const manifest = await readManifestFile(manifestPath);

    validateManifestId(manifest);
    let webViewType: devSettings.WebViewType | undefined;

    if (webViewString === undefined) {
      webViewType = await devSettings.getWebView(manifest.id!);
    } else {
      webViewType = parseWebViewType(webViewString)
      await devSettings.setWebView(manifest.id!, webViewType);
    }

    const webViewTypeName = devSettings.toWebViewTypeName(webViewType);
    console.log(webViewTypeName
      ? `The web view type is set to ${webViewTypeName}.`
      : "The web view type has not been set.");
      usageDataObject.reportSuccess("webView");
  } catch (err) {
    usageDataObject.reportException("webView", err);
    logErrorMessage(err);
  }
}
