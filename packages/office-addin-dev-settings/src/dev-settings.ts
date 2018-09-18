#!/usr/bin/env node

// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as commander from "commander";
import * as commands from "./commands";
import * as registry from "./dev-settings-registry";

export enum DebuggingMethod {
    Direct,
    Web,
}

export function clearDevSettings(addinId: string): void {
  switch (process.platform) {
    case "win32":
      return registry.clearDevSettings(addinId);
    default:
      throw new Error(`Platform not supported: ${process.platform}.`);
  }
}

export function configureSourceBundleUrl(addinId: string, host: string, port: string, path: string, extension: string): void {
  switch (process.platform) {
    case "win32":
      return registry.configureSourceBundleUrl(addinId, host, port, path, extension);
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
      return registry.enableDebugging(addinId, enable, method);
    default:
      throw new Error(`Platform not supported: ${process.platform}.`);
  }
}

export function enableLiveReload(addinId: string, enable: boolean = true): void {
  switch (process.platform) {
    case "win32":
      return registry.enableLiveReload(addinId, enable);
    default:
      throw new Error(`Platform not supported: ${process.platform}.`);
  }
}

commander
    .command("clear [manifestPath]")
    .action(commands.clear);

commander
    .command("disable-debugging [manifestPath]")
    .action(commands.disableDebugging);

commander
    .command("disable-live-reload [manifestPath]")
    .action(commands.disableLiveReload);

commander
    .command("enable-debugging [manifestPath]")
    .option("--debug-method <method>", "Specify the debug method: 'direct' or 'web'.")
    .action(commands.enableDebugging);

commander
    .command("enable-live-reload [manifestPath]")
    .action(commands.enableLiveReload);

commander.parse(process.argv);
