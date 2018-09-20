#!/usr/bin/env node

// copyright (c) Microsoft Corporation. All rights reserved.
// licensed under the MIT license.

import { DebuggingMethod } from "./dev-settings";
import * as registry from "./registry";

const DeveloperSettingsRegistryKey: string = `HKEY_CURRENT_USER\\SOFTWARE\\Microsoft\\Office\\16.0\\Wef\\Developer`;

export async function clearDevSettings(addinId: string): Promise<void> {
  return deleteDeveloperSettingsRegistryKey(addinId);
}

export async function configureSourceBundleUrl(addinId: string, host?: string, port?: string, path?: string, extension?: string): Promise<void> {
  const key: string = getDeveloperSettingsRegistryKey(addinId);

  if (host) {
    await registry.addStringValue(key, "SourceBundleHost", host);
  }

  if (port) {
    await registry.addStringValue(key, "SourceBundlePort", port);
  }

  if (path) {
    await registry.addStringValue(key, "SourceBundlePath", path);

    // for now, include old name
    await registry.addStringValue(key, "DebugBundlePath", path);
  }

  if (extension) {
    await registry.addStringValue(key, "SourceBundleExtension", extension);
  }
}

export async function deleteDeveloperSettingsRegistryKey(addinId: string): Promise<void> {
  const key: string = getDeveloperSettingsRegistryKey(addinId);
  return registry.deleteKey(key);
}

export async function enableDebugging(addinId: string, enable: boolean = true, method: DebuggingMethod = DebuggingMethod.Web): Promise<void> {
  const key: string = getDeveloperSettingsRegistryKey(addinId);
  const useDirectDebugger: boolean = enable && (method === DebuggingMethod.Direct);
  const useWebDebugger: boolean = enable && (method === DebuggingMethod.Web);

  await registry.addBooleanValue(key, "UseDirectDebugger", useDirectDebugger);
  await registry.addBooleanValue(key, "UseWebDebugger", useWebDebugger);

  // for now, include old name
  await registry.addBooleanValue(key, "Debugging", useWebDebugger);
}

export async function enableLiveReload(addinId: string, enable: boolean = true): Promise<void> {
  const key: string = getDeveloperSettingsRegistryKey(addinId);
  return registry.addBooleanValue(key, "UseLiveReload", enable);
}

export function getDeveloperSettingsRegistryKey(addinId: string): string {
  if (!addinId) { throw new Error("addinId is required."); }
  if (typeof addinId !== "string") { throw new Error("addinId should be a string."); }
  return `${DeveloperSettingsRegistryKey}\\${addinId}`;
}
