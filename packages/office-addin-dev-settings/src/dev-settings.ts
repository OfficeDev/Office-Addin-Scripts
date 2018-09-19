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

export function configureSourceBundleUrl(addinId: string, host?: string, port?: string, path?: string, extension?: string): void {
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

if (process.argv[1].endsWith("\\dev-settings.js")) {
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

  commander
    .command("source-bundle-url <manifestPath>")
    .description("Specify values for components of the url used to obtain the source bundle.")
    .option("-h,--host <host>", "Specify the host name to use (instead of 'localhost').")
    .option("-p,--port <port>", "Specify the port number to use (instead of 9229).")
    .option("--path <path>", "Specify the path to use.")
    .option("-e,--extension <extension>", `Specify the extension to use (instead of ".bundle").`)
    .action(commands.configureSourceBundleUrl);

  commander.parse(process.argv);
}
