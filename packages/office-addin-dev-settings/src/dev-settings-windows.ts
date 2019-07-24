#!/usr/bin/env node

// copyright (c) Microsoft Corporation. All rights reserved.
// licensed under the MIT license.

import { DebuggingMethod, RegisteredAddin, SourceBundleUrlComponents } from "./dev-settings";
import * as registry from "./registry";

const DeveloperSettingsRegistryKey: string = `HKEY_CURRENT_USER\\SOFTWARE\\Microsoft\\Office\\16.0\\Wef\\Developer`;

const RuntimeLogging: string = "RuntimeLogging";
const SourceBundleExtension: string = "SourceBundleExtension";
const SourceBundleHost: string = "SourceBundleHost";
const SourceBundlePath: string = "SourceBundlePath";
const SourceBundlePort: string = "SourceBundlePort";
const UseDirectDebugger: string = "UseDirectDebugger";
const UseLiveReload: string = "UseLiveReload";
const UseProxyDebugger: string = "UseWebDebugger";

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

export async function enableDebugging(addinId: string, enable: boolean = true, method: DebuggingMethod = DebuggingMethod.Proxy): Promise<void> {
  const key = getDeveloperSettingsRegistryKey(addinId);
  const useDirectDebugger: boolean = enable && (method === DebuggingMethod.Direct);
  const useProxyDebugger: boolean = enable && (method === DebuggingMethod.Proxy);

  await registry.addBooleanValue(key, UseDirectDebugger, useDirectDebugger);
  await registry.addBooleanValue(key, UseProxyDebugger, useProxyDebugger);
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
  if (!addinId) { throw new Error("addinId is required."); }
  if (typeof addinId !== "string") { throw new Error("addinId should be a string."); }
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

export async function getRegisteredAddIns(): Promise<RegisteredAddin[]> {
  const key = new registry.RegistryKey(`${DeveloperSettingsRegistryKey}`);

  const values = await registry.getValues(key);

  // if the registry value name and data are the same, then the manifest path was used as the name
  return values.map(value => new RegisteredAddin((value.name !== value.data) ? value.name : "", value.data));
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
    await registry.getStringValue(key, SourceBundleExtension),
  );
  return components;
}

export async function isDebuggingEnabled(addinId: string): Promise<boolean> {
  const key: registry.RegistryKey = getDeveloperSettingsRegistryKey(addinId);

  const useDirectDebugger: boolean = isRegistryValueTrue(await registry.getValue(key, UseDirectDebugger));
  const useWebDebugger: boolean = isRegistryValueTrue(await registry.getValue(key, UseProxyDebugger));

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

export async function registerAddIn(addinId: string, manifestPath: string) {
  const key = new registry.RegistryKey(`${DeveloperSettingsRegistryKey}`);

  await registry.deleteValue(key, manifestPath); // in case the manifest path was previously used as the key
  return registry.addStringValue(key, addinId, manifestPath);
}

export async function setSourceBundleUrl(addinId: string, components: SourceBundleUrlComponents): Promise<void> {
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

export async function unregisterAddIn(addinId: string, manifestPath: string): Promise<void> {
  const key = new registry.RegistryKey(`${DeveloperSettingsRegistryKey}`);

  if (addinId) {
    await registry.deleteValue(key, addinId);
  }

  // since the key used to be manifest path, delete it too
  if (manifestPath) {
    await registry.deleteValue(key, manifestPath);
  }
}

export async function unregisterAllAddIns(): Promise<void> {
  const key = new registry.RegistryKey(`${DeveloperSettingsRegistryKey}`);
  const values = await registry.getValues(key);

  values.forEach(async value => await registry.deleteValue(key, value.name));
}
