#!/usr/bin/env node

// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import { platform } from "os";
import * as registry from "./dev-settings-registry";

export enum DebuggingMethod {
    Direct,
    Web,
}

export function clearDevSettings(addinId: string): void {
  switch (process.platform) {
    case "win32":
      registry.clearDevSettings(addinId);
    default:
      throw new Error(`Platform not supported: ${process.platform}.`);
  }
}

export function configureSourceBundleUrl(addinId: string, host: string, port: string, path: string, extension: string): void {
  switch (process.platform) {
    case "win32":
      registry.configureSourceBundleUrl(addinId, host, port, path, extension);
    default:
      throw new Error(`Platform not supported: ${process.platform}.`);
  }
}

export function disableDebugging(addinId: string): void {
    enableDebugging(addinId, false);
}

export function disableLiveReload(addinId: string): void {
  enableLiveReload(addinId, false);
}

export function enableDebugging(addinId: string, enable: boolean = true, method: DebuggingMethod = DebuggingMethod.Web): void {
  switch (process.platform) {
    case "win32":
      registry.enableDebugging(addinId, enable, method);
    default:
      throw new Error(`Platform not supported: ${process.platform}.`);
  }
}

export function enableLiveReload(addinId: string, enable: boolean = true): void {
  switch (process.platform) {
    case "win32":
      registry.enableLiveReload(addinId, enable);
    default:
      throw new Error(`Platform not supported: ${process.platform}.`);
  }
}
