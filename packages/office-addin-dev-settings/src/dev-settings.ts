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

export async function clearDevSettings(addinId: string): Promise<void> {
  switch (process.platform) {
    case "win32":
      return registry.clearDevSettings(addinId);
  default:
    throw new Error(`Platform not supported: ${process.platform}.`);
  }
}

export async function configureSourceBundleUrl(addinId: string, host?: string, port?: string, path?: string, extension?: string): Promise<void> {
  switch (process.platform) {
    case "win32":
      return registry.configureSourceBundleUrl(addinId, host, port, path, extension);
    default:
      throw new Error(`Platform not supported: ${process.platform}.`);
  }
}

export async function disableDebugging(addinId: string): Promise<void> {
  return enableDebugging(addinId, false);
}

export async function disableLiveReload(addinId: string): Promise<void> {
  return enableLiveReload(addinId, false);
}

export async function enableDebugging(addinId: string, enable: boolean = true, method: DebuggingMethod = DebuggingMethod.Web): Promise<void> {
  switch (process.platform) {
    case "win32":
      return registry.enableDebugging(addinId, enable, method);
    default:
      throw new Error(`Platform not supported: ${process.platform}.`);
  }
}

export async function enableLiveReload(addinId: string, enable: boolean = true): Promise<void> {
  switch (process.platform) {
    case "win32":
      return registry.enableLiveReload(addinId, enable);
  default:
    throw new Error(`Platform not supported: ${process.platform}.`);
  }
}

export async function isDebuggingEnabled(addinId: string): Promise<boolean> {
  switch (process.platform) {
    case "win32":
      return registry.isDebuggingEnabled(addinId);
    default:
      throw new Error(`Platform not supported: ${process.platform}.`);
  }
}

export async function isLiveReloadEnabled(addinId: string): Promise<boolean> {
  switch (process.platform) {
    case "win32":
      return registry.isLiveReloadEnabled(addinId);
    default:
      throw new Error(`Platform not supported: ${process.platform}.`);
  }
}
if (process.argv[1].endsWith("\\dev-settings.js")) {
  commander
    .command("clear [manifestPath]")
    .description("Clear all dev settings for the add-in.")
    .action(commands.clear);

  commander
    .command("disable-debugging [manifestPath]")
    .description("Disable debugging of the add-in.")
    .action(commands.disableDebugging);

  commander
    .command("disable-live-reload [manifestPath]")
    .description("Disable live reload for the add-in.")
    .action(commands.disableLiveReload);

  commander
    .command("enable-debugging [manifestPath]")
    .description("Enable debugging for the add-in.")
    .option("--debug-method <method>", "Specify the debug method: 'direct' or 'web'.")
    .action(commands.enableDebugging);

  commander
    .command("enable-live-reload [manifestPath]")
    .description("Enable live reload for the add-in.")
    .action(commands.enableLiveReload);

  commander
    .command("is-debugging-enabled [manifestPath]")
    .description("Display whether debugging is enabled for the add-in.")
    .action(commands.isDebuggingEnabled);

  commander
    .command("is-live-reload-enabled [manifestPath]")
    .description("Display whether live reload is enabled.")
    .action(commands.isDebuggingEnabled);

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
