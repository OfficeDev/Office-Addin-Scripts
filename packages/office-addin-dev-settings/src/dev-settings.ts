#!/usr/bin/env node

// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as registry from './registry';

export enum DebuggingMethod {
    Direct,
    Web
}

export function clearDevSettings(addinId: string) {
  registry.deleteDeveloperSettingsRegistryKey(addinId);
}

export function disableDebugging(addinId: string) {
    enableDebugging(addinId, false);
}

export function enableDebugging(addinId: string, enable: boolean = true, method: DebuggingMethod = DebuggingMethod.Web) {
  const key = registry.getDeveloperSettingsRegistryKey(addinId);
  const useDirectDebugger = enable && (method == DebuggingMethod.Direct);
  const useWebDebugger = enable && (method == DebuggingMethod.Web);

  registry.addBooleanValue(key, 'UseDirectDebugger', useDirectDebugger);
  registry.addBooleanValue(key, 'UseWebDebugger', useWebDebugger);

  // for now, include old name
  registry.addBooleanValue(key, 'EnableDebugging', useWebDebugger);
}

export function disableLiveReload(addinId: string) {
    enableLiveReload(addinId, false);
}

export function enableLiveReload(addinId: string, enable: boolean = true) {
  const key = registry.getDeveloperSettingsRegistryKey(addinId);
  registry.addBooleanValue(key, 'UseLiveReload', enable);
}

export function configureSourceBundleUrl(addinId: string, host: string, port: string, path: string, extension: string) {
  const key = registry.getDeveloperSettingsRegistryKey(addinId);

  if (host) {
    registry.addStringValue(key, 'SourceBundleHost', host);
  }

  if (port) {
    registry.addStringValue(key, 'SourceBundlePort', port);
  }

  if (path) {
    registry.addStringValue(key, 'SourceBundlePath', path);

    // for now, include old name
    registry.addStringValue(key, 'DebugBundlePath', path);
  }
  
  if (extension) {
    registry.addStringValue(key, 'SourceBundleExtension', extension);
  }
}

