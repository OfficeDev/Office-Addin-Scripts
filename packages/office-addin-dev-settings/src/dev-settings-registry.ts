#!/usr/bin/env node

// copyright (c) Microsoft Corporation. All rights reserved.
// licensed under the MIT license.

import { DebuggingMethod } from "./dev-settings";
import * as registry from "./registry";

const DeveloperSettingsRegistryKey: string = `HKEY_CURRENT_USER\\SOFTWARE\\Microsoft\\Office\\16.0\\Wef\\Developer`;

export function clearDevSettings(addinId: string): void {
  deleteDeveloperSettingsRegistryKey(addinId);
}

export function configureSourceBundleUrl(addinId: string, host?: string, port?: string, path?: string, extension?: string): void {
  const key: string = getDeveloperSettingsRegistryKey(addinId);

  if (host) {
    registry.addStringValue(key, "SourceBundleHost", host);
  }

  if (port) {
    registry.addStringValue(key, "SourceBundlePort", port);
  }

  if (path) {
    registry.addStringValue(key, "SourceBundlePath", path);

    // for now, include old name
    registry.addStringValue(key, "DebugBundlePath", path);
  }

  if (extension) {
    registry.addStringValue(key, "SourceBundleExtension", extension);
  }
}

export function deleteDeveloperSettingsRegistryKey(addinId: string): void {
  const key: string = getDeveloperSettingsRegistryKey(addinId);
  registry.deleteKey(key);
}

export function enableDebugging(addinId: string, enable: boolean = true, method: DebuggingMethod = DebuggingMethod.Web): void {
  const key: string = getDeveloperSettingsRegistryKey(addinId);
  const useDirectDebugger: boolean = enable && (method === DebuggingMethod.Direct);
  const useWebDebugger: boolean = enable && (method === DebuggingMethod.Web);

  registry.addBooleanValue(key, "UseDirectDebugger", useDirectDebugger);
  registry.addBooleanValue(key, "UseWebDebugger", useWebDebugger);

  // for now, include old name
  registry.addBooleanValue(key, "EnableDebugging", useWebDebugger);
}

export function enableLiveReload(addinId: string, enable: boolean = true): void {
  const key: string = getDeveloperSettingsRegistryKey(addinId);
  registry.addBooleanValue(key, "UseLiveReload", enable);
}

export function getDeveloperSettingsRegistryKey(addinId: string): string {
  if (!addinId) { throw new Error("addinId is required."); }
  if (typeof addinId !== "string") { throw new Error("addinId should be a string."); }
  return `${DeveloperSettingsRegistryKey}\\${addinId}`;
}
