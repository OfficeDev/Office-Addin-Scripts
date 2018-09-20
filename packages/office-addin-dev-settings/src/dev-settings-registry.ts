#!/usr/bin/env node

// copyright (c) Microsoft Corporation. All rights reserved.
// licensed under the MIT license.

import { DebuggingMethod, SourceBundleUrlComponents } from "./dev-settings";
import * as registry from "./registry";

const DeveloperSettingsRegistryKey: string = `HKEY_CURRENT_USER\\SOFTWARE\\Microsoft\\Office\\16.0\\Wef\\Developer`;

export async function clearDevSettings(addinId: string): Promise<void> {
  return deleteDeveloperSettingsRegistryKey(addinId);
}

export async function deleteDeveloperSettingsRegistryKey(addinId: string): Promise<void> {
  const key = getDeveloperSettingsRegistryKey(addinId);
  return registry.deleteKey(key);
}

export async function enableDebugging(addinId: string, enable: boolean = true, method: DebuggingMethod = DebuggingMethod.Web): Promise<void> {
  const key = getDeveloperSettingsRegistryKey(addinId);
  const useDirectDebugger: boolean = enable && (method === DebuggingMethod.Direct);
  const useWebDebugger: boolean = enable && (method === DebuggingMethod.Web);

  await registry.addBooleanValue(key, "UseDirectDebugger", useDirectDebugger);
  await registry.addBooleanValue(key, "UseWebDebugger", useWebDebugger);

  // for now, include old name
  await registry.addBooleanValue(key, "Debugging", useWebDebugger);
}

export async function enableLiveReload(addinId: string, enable: boolean = true): Promise<void> {
  const key = getDeveloperSettingsRegistryKey(addinId);
  return registry.addBooleanValue(key, "UseLiveReload", enable);
}

export function getDeveloperSettingsRegistryKey(addinId: string): registry.RegistryKey {
  if (!addinId) { throw new Error("addinId is required."); }
  if (typeof addinId !== "string") { throw new Error("addinId should be a string."); }
  return new registry.RegistryKey(`${DeveloperSettingsRegistryKey}\\${addinId}`);
}

export async function getSourceBundleUrl(addinId: string): Promise<SourceBundleUrlComponents> {
  const key = getDeveloperSettingsRegistryKey(addinId);
  const components = new SourceBundleUrlComponents(
    await registry.getStringValue(key, "SourceBundleHost"),
    await registry.getStringValue(key, "SourceBundlePort"),
    await registry.getStringValue(key, "SourceBundlePath"),
    await registry.getStringValue(key, "SourceBundleExtension"),
  );
  return components;
}

export async function isDebuggingEnabled(addinId: string): Promise<boolean> {
  const key: registry.RegistryKey = getDeveloperSettingsRegistryKey(addinId);

  const useDirectDebugger: boolean = isRegistryValueTrue(await registry.getValue(key, "UseDirectDebugger"));
  const useWebDebugger: boolean = isRegistryValueTrue(await registry.getValue(key, "UseWebDebugger"))
    ||  isRegistryValueTrue(await registry.getValue(key, "Debugging")); // for now, include old name

  return useDirectDebugger || useWebDebugger;
}

export async function isLiveReloadEnabled(addinId: string): Promise<boolean> {
  const key = getDeveloperSettingsRegistryKey(addinId);

  const enabled: boolean = isRegistryValueTrue(await registry.getValue(key, "UseLiveReload"));

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

export async function setSourceBundleUrl(addinId: string, components: SourceBundleUrlComponents): Promise<void> {
  const key = getDeveloperSettingsRegistryKey(addinId);

  if (components.host) {
    await registry.addStringValue(key, "SourceBundleHost", components.host);
  } else {
    await registry.deleteValue(key, "SourceBundleHost");
  }

  if (components.port) {
    await registry.addStringValue(key, "SourceBundlePort", components.port);
  } else {
    await registry.deleteValue(key, "SourceBundlePort");
  }

  if (components.path) {
    await registry.addStringValue(key, "SourceBundlePath", components.path);

    // for now, include old name
    await registry.addStringValue(key, "DebugBundlePath", components.path);
  } else {
    await registry.deleteValue(key, "SourceBundlePath");

    // for now include old name
    await registry.deleteValue(key, "DebugBundlePath");
  }

  if (components.extension !== undefined) {
    await registry.addStringValue(key, "SourceBundleExtension", components.extension);
  } else {
    await registry.deleteValue(key, "SourceBundleExtension");
  }
}
