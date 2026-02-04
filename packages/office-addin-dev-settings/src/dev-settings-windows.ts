#!/usr/bin/env node

// copyright (c) Microsoft Corporation. All rights reserved.
// licensed under the MIT license.

import {
  exportMetadataPackage,
  getOfficeAppsForManifestHosts,
  ManifestInfo,
  OfficeApp,
  OfficeAddinManifest,
  ManifestType,
} from "office-addin-manifest";
import {
  DebuggingMethod,
  RegisteredAddin,
  SourceBundleUrlComponents,
  WebViewType,
} from "./dev-settings";
import { ExpectedError } from "office-addin-usage-data";
import * as registry from "./registry";
import { registerWithTeams, uninstallWithTeams } from "./publish";
import fspath from "path";

/* global process */

const DeveloperSettingsRegistryKey: string = `HKEY_CURRENT_USER\\SOFTWARE\\Microsoft\\Office\\16.0\\Wef\\Developer`;

const OpenDevTools: string = "OpenDevTools";
export const OutlookSideloadManifestPath: string = "OutlookSideloadManifestPath";
const RefreshAddins: string = "RefreshAddins";
const RuntimeLogging: string = "RuntimeLogging";
const SourceBundleExtension: string = "SourceBundleExtension";
const SourceBundleHost: string = "SourceBundleHost";
const SourceBundlePath: string = "SourceBundlePath";
const SourceBundlePort: string = "SourceBundlePort";
const TitleId: string = "TitleId";
const UseDirectDebugger: string = "UseDirectDebugger";
const UseLiveReload: string = "UseLiveReload";
const UseProxyDebugger: string = "UseWebDebugger";
const WebViewSelection: string = "WebViewSelection";

export async function clearDevSettings(addinId: string): Promise<void> {
  return deleteDeveloperSettingsRegistryKey(addinId);
}

export async function deleteDeveloperSettingsRegistryKey(addinId: string): Promise<void> {
  const key = getDeveloperSettingsRegistryKey(addinId);
  return registry.deleteKey(key);
}

export async function disableRuntimeLogging() {
  const key = getDeveloperSettingsRegistryKey(RuntimeLogging);

  return registry.deleteKey(key);
}

export async function enableDebugging(
  addinId: string,
  enable: boolean = true,
  method: DebuggingMethod = DebuggingMethod.Proxy,
  openDevTools: boolean = false
): Promise<void> {
  const key = getDeveloperSettingsRegistryKey(addinId);
  const useDirectDebugger: boolean = enable && method === DebuggingMethod.Direct;
  const useProxyDebugger: boolean = enable && method === DebuggingMethod.Proxy;

  await registry.addBooleanValue(key, UseDirectDebugger, useDirectDebugger);
  await registry.addBooleanValue(key, UseProxyDebugger, useProxyDebugger);

  if (enable && openDevTools) {
    await registry.addBooleanValue(key, OpenDevTools, true);
  } else {
    await registry.deleteValue(key, OpenDevTools);
  }
}

export async function enableLiveReload(addinId: string, enable: boolean = true): Promise<void> {
  const key = getDeveloperSettingsRegistryKey(addinId);
  return registry.addBooleanValue(key, UseLiveReload, enable);
}

export async function enableRuntimeLogging(path: string): Promise<void> {
  const key = getDeveloperSettingsRegistryKey(RuntimeLogging);

  return registry.addStringValue(key, "", path); // empty string for the default value
}

export function getDeveloperSettingsRegistryKey(addinId: string): registry.RegistryKey {
  if (!addinId) {
    throw new ExpectedError("The addIn parameter is required.");
  }
  if (typeof addinId !== "string") {
    throw new ExpectedError("The addIn parameter should be a string.");
  }
  return new registry.RegistryKey(`${DeveloperSettingsRegistryKey}\\${addinId}`);
}

export async function getEnabledDebuggingMethods(addinId: string): Promise<DebuggingMethod[]> {
  const key: registry.RegistryKey = getDeveloperSettingsRegistryKey(addinId);
  const methods: DebuggingMethod[] = [];

  if (isRegistryValueTrue(await registry.getValue(key, UseDirectDebugger))) {
    methods.push(DebuggingMethod.Direct);
  }

  if (isRegistryValueTrue(await registry.getValue(key, UseProxyDebugger))) {
    methods.push(DebuggingMethod.Proxy);
  }

  return methods;
}

export async function getOpenDevTools(addinId: string): Promise<boolean> {
  const key: registry.RegistryKey = getDeveloperSettingsRegistryKey(addinId);

  return isRegistryValueTrue(await registry.getValue(key, OpenDevTools));
}

export async function getRegisteredAddIns(): Promise<RegisteredAddin[]> {
  const key = new registry.RegistryKey(`${DeveloperSettingsRegistryKey}`);
  const values = await registry.getValues(key);
  const filteredValues = values.filter((value) => value.name !== RefreshAddins);

  // if the registry value name and data are the same, then the manifest path was used as the name
  return filteredValues.map(
    (value) => new RegisteredAddin(value.name !== value.data ? value.name : "", value.data)
  );
}

export async function getRuntimeLoggingPath(): Promise<string | undefined> {
  const key = getDeveloperSettingsRegistryKey(RuntimeLogging);

  return registry.getStringValue(key, ""); // empty string for the default value
}

export async function getSourceBundleUrl(addinId: string): Promise<SourceBundleUrlComponents> {
  const key = getDeveloperSettingsRegistryKey(addinId);
  const components = new SourceBundleUrlComponents(
    await registry.getStringValue(key, SourceBundleHost),
    await registry.getStringValue(key, SourceBundlePort),
    await registry.getStringValue(key, SourceBundlePath),
    await registry.getStringValue(key, SourceBundleExtension)
  );
  return components;
}

export async function getWebView(addinId: string): Promise<WebViewType | undefined> {
  const key = getDeveloperSettingsRegistryKey(addinId);
  const webViewString = await registry.getStringValue(key, WebViewSelection);
  return toWebViewType(webViewString);
}

export async function isDebuggingEnabled(addinId: string): Promise<boolean> {
  const key: registry.RegistryKey = getDeveloperSettingsRegistryKey(addinId);

  const useDirectDebugger: boolean = isRegistryValueTrue(
    await registry.getValue(key, UseDirectDebugger)
  );
  const useWebDebugger: boolean = isRegistryValueTrue(
    await registry.getValue(key, UseProxyDebugger)
  );

  return useDirectDebugger || useWebDebugger;
}

export async function isLiveReloadEnabled(addinId: string): Promise<boolean> {
  const key = getDeveloperSettingsRegistryKey(addinId);

  const enabled: boolean = isRegistryValueTrue(await registry.getValue(key, UseLiveReload));

  return enabled;
}

function isRegistryValueTrue(value?: registry.RegistryValue): boolean {
  if (value) {
    switch (value.type) {
      case registry.RegistryTypes.REG_DWORD:
      case registry.RegistryTypes.REG_QWORD:
      case registry.RegistryTypes.REG_SZ:
        return parseInt(value.data, undefined) !== 0;
    }
  }

  return false;
}

export async function registerAddIn(manifestPath: string, registration?: string): Promise<void> {
  const manifest: ManifestInfo = await OfficeAddinManifest.readManifestFile(manifestPath);
  const appsInManifest = getOfficeAppsForManifestHosts(manifest.hosts);

  // Register using the service
  if (
    manifest.manifestType === ManifestType.JSON ||
    appsInManifest.indexOf(OfficeApp.Outlook) >= 0
  ) {
    if (!registration) {
      let filePath = "";
      if (manifest.manifestType === ManifestType.JSON) {
        const targetPath: string = fspath.join(process.env.TEMP as string, "manifest.zip");
        filePath = await exportMetadataPackage(targetPath, manifestPath);
      } else if (manifest.manifestType === ManifestType.XML) {
        filePath = manifestPath;
      }
      registration = await registerWithTeams(filePath);
      await enableRefreshAddins();
    }

    const key = getDeveloperSettingsRegistryKey(OutlookSideloadManifestPath);
    await registry.addStringValue(key, TitleId, registration);
  }

  const key = new registry.RegistryKey(`${DeveloperSettingsRegistryKey}`);

  await registry.deleteValue(key, manifestPath); // in case the manifest path was previously used as the key
  return registry.addStringValue(key, manifest.id || "", manifestPath);
}

export async function setSourceBundleUrl(
  addinId: string,
  components: SourceBundleUrlComponents
): Promise<void> {
  const key = getDeveloperSettingsRegistryKey(addinId);

  if (components.host !== undefined) {
    if (components.host) {
      await registry.addStringValue(key, SourceBundleHost, components.host);
    } else {
      await registry.deleteValue(key, SourceBundleHost);
    }
  }

  if (components.port !== undefined) {
    if (components.port) {
      await registry.addStringValue(key, SourceBundlePort, components.port);
    } else {
      await registry.deleteValue(key, SourceBundlePort);
    }
  }

  if (components.path !== undefined) {
    if (components.path) {
      await registry.addStringValue(key, SourceBundlePath, components.path);
    } else {
      await registry.deleteValue(key, SourceBundlePath);
    }
  }

  if (components.extension !== undefined) {
    if (components.extension) {
      await registry.addStringValue(key, SourceBundleExtension, components.extension);
    } else {
      await registry.deleteValue(key, SourceBundleExtension);
    }
  }
}

export async function setWebView(
  addinId: string,
  webViewType: WebViewType | undefined
): Promise<void> {
  const key = getDeveloperSettingsRegistryKey(addinId);
  switch (webViewType) {
    case undefined:
    case WebViewType.Default:
      await registry.deleteValue(key, WebViewSelection);
      break;
    case WebViewType.EdgeChromium: {
      const webViewString: string = webViewType as string;
      await registry.addStringValue(key, WebViewSelection, webViewString);
      break;
    }
    default:
      throw new ExpectedError(`The webViewType ${webViewType} is not supported.`);
  }
}

function toWebViewType(webViewString?: string): WebViewType | undefined {
  switch (webViewString ? webViewString.toLowerCase() : undefined) {
    case "edge chromium":
      return WebViewType.EdgeChromium;
    default:
      return undefined;
  }
}

export function toWebViewTypeName(webViewType?: WebViewType): string | undefined {
  switch (webViewType) {
    case WebViewType.EdgeChromium:
      return "Microsoft Edge (Chromium)";
    default:
      return undefined;
  }
}

export async function unregisterAddIn(addinId: string, manifestPath: string): Promise<void> {
  const key = new registry.RegistryKey(`${DeveloperSettingsRegistryKey}`);

  if (addinId) {
    const manifest: ManifestInfo = await OfficeAddinManifest.readManifestFile(manifestPath);
    const appsInManifest = getOfficeAppsForManifestHosts(manifest.hosts);

    // Unregister using the service
    // TODO: remove if statement when the service supports all hosts
    if (
      manifest.manifestType === ManifestType.JSON ||
      appsInManifest.indexOf(OfficeApp.Outlook) >= 0
    ) {
      await unacquire(key, addinId);
    }
    await registry.deleteValue(key, addinId);
  }

  // since the key used to be manifest path, delete it too
  if (manifestPath) {
    await registry.deleteValue(key, manifestPath);
  }
}

export async function unregisterAllAddIns(): Promise<void> {
    const registeredAddins: RegisteredAddin[] = await getRegisteredAddIns();

    for (const addin of registeredAddins) {
        await unregisterAddIn(addin.id, addin.manifestPath);
    }
}

async function unacquire(key: registry.RegistryKey, id: string) {
  const manifest = await registry.getStringValue(key, id);
  if (manifest != undefined) {
    const key = getDeveloperSettingsRegistryKey(OutlookSideloadManifestPath);
    const registration = await registry.getStringValue(key, TitleId);
    // Try manifest id first, fall back to titleId from registry
    const uninstallId = id || registration;
    if (uninstallId && await uninstallWithTeams(uninstallId)) {
      if (registration != undefined) {
        registry.deleteValue(key, TitleId);
      }
      await enableRefreshAddins();
    }
  }
}

async function enableRefreshAddins() {
  const key = new registry.RegistryKey(`${DeveloperSettingsRegistryKey}`);
  await registry.addBooleanValue(key, RefreshAddins, true);
}

export async function disableRefreshAddins() {
  const key = new registry.RegistryKey(`${DeveloperSettingsRegistryKey}`);
  await registry.deleteValue(key, RefreshAddins);
}
