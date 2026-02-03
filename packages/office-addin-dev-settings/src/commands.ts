// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import { OptionValues } from "commander";
import { logErrorMessage } from "office-addin-usage-data";
import {
  ManifestInfo,
  OfficeApp,
  parseOfficeApp,
  OfficeAddinManifest,
} from "office-addin-manifest";
import { AppType, parseAppType } from "./appType";
import * as devSettings from "./dev-settings";
import { sideloadAddIn } from "./sideload";
import { usageDataObject } from "./defaults";
import { ExpectedError } from "office-addin-usage-data";
import { AccountOperation, updateM365Account } from "./publish";

/* global console process */


export async function clear(manifestPath: string) {
  try {
    const manifest = await OfficeAddinManifest.readManifestFile(manifestPath);

    validateManifestId(manifest);

    await devSettings.clearDevSettings(manifest.id!);

    console.log("Developer settings have been cleared.");
    usageDataObject.reportSuccess("clear");
  } catch (err: any) {
    usageDataObject.reportException("clear", err);
    logErrorMessage(err);
  }
}

export async function debugging(manifestPath: string, options: OptionValues) {
  try {
    if (options.enable) {
      await enableDebugging(manifestPath, options);
    } else if (options.disable) {
      await disableDebugging(manifestPath);
    } else {
      await isDebuggingEnabled(manifestPath);
    }
    usageDataObject.reportSuccess("debugging");
  } catch (err: any) {
    usageDataObject.reportException("debugging", err);
    logErrorMessage(err);
  }
}

function displaySourceBundleUrl(components: devSettings.SourceBundleUrlComponents) {
  console.log(
    `host: ${components.host !== undefined ? `"${components.host}"` : '"localhost" (default)'}`
  );
  console.log(
    `port: ${components.port !== undefined ? `"${components.port}"` : '"8081" (default)'}`
  );
  console.log(`path: ${components.path !== undefined ? `"${components.path}"` : "(default)"}`);
  console.log(
    `extension: ${components.extension !== undefined ? `"${components.extension}"` : '".bundle" (default)'}`
  );
  console.log();
  console.log(`Source bundle url: ${components.url}`);
  console.log();
}

export async function disableDebugging(manifestPath: string) {
  try {
    const manifest = await OfficeAddinManifest.readManifestFile(manifestPath);

    validateManifestId(manifest);

    await devSettings.disableDebugging(manifest.id!);

    console.log("Debugging has been disabled.");
    usageDataObject.reportSuccess("disableDebugging()");
  } catch (err: any) {
    usageDataObject.reportException("disableDebugging()", err);
    logErrorMessage(err);
  }
}

export async function disableLiveReload(manifestPath: string) {
  try {
    const manifest = await OfficeAddinManifest.readManifestFile(manifestPath);

    validateManifestId(manifest);

    await devSettings.disableLiveReload(manifest.id!);

    console.log("Live reload has been disabled.");
    usageDataObject.reportSuccess("disableLiveReload()");
  } catch (err: any) {
    usageDataObject.reportException("disableLiveReload()", err);
    logErrorMessage(err);
  }
}

export async function disableRuntimeLogging() {
  try {
    await devSettings.disableRuntimeLogging();

    console.log("Runtime logging has been disabled.");
    usageDataObject.reportSuccess("disableRuntimeLogging()");
  } catch (err: any) {
    usageDataObject.reportException("disableRuntimeLogging()", err);
    logErrorMessage(err);
  }
}

export async function enableDebugging(manifestPath: string, options: OptionValues) {
  try {
    const manifest = await OfficeAddinManifest.readManifestFile(manifestPath);

    validateManifestId(manifest);

    await devSettings.enableDebugging(
      manifest.id!,
      true,
      toDebuggingMethod(options.debugMethod),
      options.openDevTools
    );

    console.log("Debugging has been enabled.");
    usageDataObject.reportSuccess("enableDebugging()");
  } catch (err: any) {
    usageDataObject.reportException("enableDebugging()", err);
    logErrorMessage(err);
  }
}

export async function enableLiveReload(manifestPath: string) {
  try {
    const manifest = await OfficeAddinManifest.readManifestFile(manifestPath);

    validateManifestId(manifest);

    await devSettings.enableLiveReload(manifest.id!);

    console.log("Live reload has been enabled.");
    usageDataObject.reportSuccess("enableLiveReload()");
  } catch (err: any) {
    usageDataObject.reportException("enableLiveReload()", err);
    logErrorMessage(err);
  }
}

export async function enableRuntimeLogging(path?: string) {
  try {
    const logPath = await devSettings.enableRuntimeLogging(path);

    console.log(`Runtime logging has been enabled. File: ${logPath}`);
    usageDataObject.reportSuccess("enableRuntimeLogging()");
  } catch (err: any) {
    usageDataObject.reportException("enableRuntimeLogging()", err);
    logErrorMessage(err);
  }
}

export async function getSourceBundleUrl(manifestPath: string) {
  try {
    const manifest = await OfficeAddinManifest.readManifestFile(manifestPath);

    validateManifestId(manifest);

    const components: devSettings.SourceBundleUrlComponents = await devSettings.getSourceBundleUrl(
      manifest.id!
    );

    displaySourceBundleUrl(components);
    usageDataObject.reportSuccess("getSourceBundleUrl()");
  } catch (err: any) {
    usageDataObject.reportException("getSourceBundleUrl()", err);
    logErrorMessage(err);
  }
}

export async function isDebuggingEnabled(manifestPath: string) {
  try {
    const manifest = await OfficeAddinManifest.readManifestFile(manifestPath);

    validateManifestId(manifest);

    const enabled: boolean = await devSettings.isDebuggingEnabled(manifest.id!);

    console.log(enabled ? "Debugging is enabled." : "Debugging is not enabled.");
    usageDataObject.reportSuccess("isDebuggingEnabled()");
  } catch (err: any) {
    usageDataObject.reportException("isDebuggingEnabled()", err);
    logErrorMessage(err);
  }
}

export async function isLiveReloadEnabled(manifestPath: string) {
  try {
    const manifest = await OfficeAddinManifest.readManifestFile(manifestPath);

    validateManifestId(manifest);

    const enabled: boolean = await devSettings.isLiveReloadEnabled(manifest.id!);

    console.log(enabled ? "Live reload is enabled." : "Live reload is not enabled.");
    usageDataObject.reportSuccess("isLiveReloadEnabled()");
  } catch (err: any) {
    usageDataObject.reportException("isLiveReloadEnabled()", err);
    logErrorMessage(err);
  }
}

export async function isRuntimeLoggingEnabled() {
  try {
    const path = await devSettings.getRuntimeLoggingPath();

    console.log(
      path ? `Runtime logging is enabled. File: ${path}` : "Runtime logging is not enabled."
    );
    usageDataObject.reportSuccess("isRuntimeLoggingEnabled()");
  } catch (err: any) {
    usageDataObject.reportException("isRuntimeLoggingEnabled()", err);
    logErrorMessage(err);
  }
}

export async function liveReload(manifestPath: string, options: OptionValues) {
  try {
    if (options.enable) {
      await enableLiveReload(manifestPath);
    } else if (options.disable) {
      await disableLiveReload(manifestPath);
    } else {
      await isLiveReloadEnabled(manifestPath);
    }
    usageDataObject.reportSuccess("liveReload");
  } catch (err: any) {
    usageDataObject.reportException("liveReload", err);
    logErrorMessage(err);
  }
}

function parseStringCommandOption(optionValue: any): string | undefined {
  return typeof optionValue === "string" ? optionValue : undefined;
}

export function parseWebViewType(webViewString?: string): devSettings.WebViewType | undefined {
  switch (webViewString ? webViewString.toLowerCase() : undefined) {
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
      throw new ExpectedError(`Please select a valid web view type instead of '${webViewString}'.`);
  }
}

export async function m365Account(
  operation: AccountOperation,
  options: OptionValues /* eslint-disable-line @typescript-eslint/no-unused-vars */
) {
  try {
    await updateM365Account(operation);
    usageDataObject.reportSuccess("m365Account");
  } catch (err: any) {
    usageDataObject.reportException("m365Account", err);
    logErrorMessage(err);
  }
}

export async function register(
  manifestPath: string,
  options: OptionValues /* eslint-disable-line @typescript-eslint/no-unused-vars */
) {
  try {
    await devSettings.registerAddIn(manifestPath);
    usageDataObject.reportSuccess("register");
  } catch (err: any) {
    usageDataObject.reportException("register", err);
    logErrorMessage(err);
  }
}

export async function registered(
  options: OptionValues /* eslint-disable-line @typescript-eslint/no-unused-vars */
) {
  try {
    const registeredAddins: devSettings.RegisteredAddin[] = await devSettings.getRegisteredAddIns();

    if (registeredAddins.length > 0) {
      for (const addin of registeredAddins) {
        let id: string = addin.id;

        if (!id && addin.manifestPath) {
          try {
            const manifest = await OfficeAddinManifest.readManifestFile(addin.manifestPath);
            id = manifest.id || "";
          } catch {
            // ignore errors
          }
        }

        console.log(`${id ? id + " " : ""}${addin.manifestPath}`);
      }
    } else {
      console.log("No add-ins are registered.");
    }
    usageDataObject.reportSuccess("registered");
  } catch (err: any) {
    usageDataObject.reportException("registered", err);
    logErrorMessage(err);
  }
}

export async function runtimeLogging(options: OptionValues) {
  try {
    if (options.enable) {
      const path: string | undefined =
        typeof options.enable === "string" ? options.enable : undefined;
      await enableRuntimeLogging(path);
    } else if (options.disable) {
      await disableRuntimeLogging();
    } else {
      await isRuntimeLoggingEnabled();
    }
    usageDataObject.reportSuccess("runtimeLogging");
  } catch (err: any) {
    usageDataObject.reportException("runtimeLogging", err);
    logErrorMessage(err);
  }
}

export async function sideload(
  manifestPath: string,
  type: string | undefined,
  options: OptionValues
) {
  try {
    const app: OfficeApp | undefined = options.app ? parseOfficeApp(options.app) : undefined;
    const canPrompt = true;
    const document: string | undefined = options.document ? options.document : undefined;
    const appType: AppType | undefined = parseAppType(
      type || process.env.npm_package_config_app_platform_to_debug
    );
    const registration: string = options.registration;

    await sideloadAddIn(manifestPath, app, canPrompt, appType, document, registration);
    usageDataObject.reportSuccess("sideload");
  } catch (err: any) {
    usageDataObject.reportException("sideload", err);
    logErrorMessage(err);
  }
}

export async function setSourceBundleUrl(manifestPath: string, options: OptionValues) {
  try {
    const manifest = await OfficeAddinManifest.readManifestFile(manifestPath);
    const host = parseStringCommandOption(options.host);
    const port = parseStringCommandOption(options.port);
    const path = parseStringCommandOption(options.path);
    const extension = parseStringCommandOption(options.extension);
    const components = new devSettings.SourceBundleUrlComponents(host, port, path, extension);

    validateManifestId(manifest);

    await devSettings.setSourceBundleUrl(manifest.id!, components);

    console.log("Configured source bundle url.");
    displaySourceBundleUrl(await devSettings.getSourceBundleUrl(manifest.id!));
    usageDataObject.reportSuccess("setSourceBundleUrl()");
  } catch (err: any) {
    usageDataObject.reportException("setSourceBundleUrl()", err);
    logErrorMessage(err);
  }
}

export async function sourceBundleUrl(manifestPath: string, options: OptionValues) {
  try {
    if (
      options.host !== undefined ||
      options.port !== undefined ||
      options.path !== undefined ||
      options.extension !== undefined
    ) {
      await setSourceBundleUrl(manifestPath, options);
    } else {
      await getSourceBundleUrl(manifestPath);
    }
    usageDataObject.reportSuccess("sourceBundleUrl");
  } catch (err: any) {
    usageDataObject.reportException("sourceBundleUrl", err);
    logErrorMessage(err);
  }
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

export async function unregister(
  manifestPath: string,
  options: OptionValues /* eslint-disable-line @typescript-eslint/no-unused-vars */
) {
  try {
    if (manifestPath === "all") {
      await devSettings.unregisterAllAddIns();
    } else {
      await devSettings.unregisterAddIn(manifestPath);
    }
    usageDataObject.reportSuccess("unregister");
  } catch (err: any) {
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
    const manifest = await OfficeAddinManifest.readManifestFile(manifestPath);

    validateManifestId(manifest);
    let webViewType: devSettings.WebViewType | undefined;

    if (webViewString === undefined) {
      webViewType = await devSettings.getWebView(manifest.id!);
    } else {
      webViewType = parseWebViewType(webViewString);
      await devSettings.setWebView(manifest.id!, webViewType);
    }

    const webViewTypeName = devSettings.toWebViewTypeName(webViewType);
    console.log(
      webViewTypeName
        ? `The web view type is set to ${webViewTypeName}.`
        : "The web view type has not been set."
    );
    usageDataObject.reportSuccess("webView");
  } catch (err: any) {
    usageDataObject.reportException("webView", err);
    logErrorMessage(err);
  }
}
