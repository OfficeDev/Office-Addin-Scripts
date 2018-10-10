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

export class SourceBundleUrlComponents {
  public host?: string;
  public port?: string;
  public path?: string;
  public extension?: string;

  public get url(): string {
    const host = (this.host !== undefined) ? this.host : "localhost";
    const port = (this.port !== undefined) ? this.port : "8081";
    const path = (this.path !== undefined) ? this.path : "{path}";
    const extension = (this.extension !== undefined) ? this.extension : ".bundle";

    return `http://${host}:${port}/${path}${extension}`;
  }

  constructor(host?: string, port?: string, path?: string, extension?: string) {
    this.host = host;
    this.port = port;
    this.path = path;
    this.extension = extension;
  }
}

export async function clearDevSettings(addinId: string): Promise<void> {
  switch (process.platform) {
    case "win32":
      return registry.clearDevSettings(addinId);
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

export async function disableRuntimeLogging(): Promise<void> {
  switch (process.platform) {
    case "win32":
      return registry.disableRuntimeLogging();
  default:
    throw new Error(`Platform not supported: ${process.platform}.`);
  }
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

export async function enableRuntimeLogging(path?: string): Promise<string> {
  switch (process.platform) {
    case "win32":
      if (!path) {
        const tempDir = process.env.TEMP;
        if (!tempDir) {
          throw new Error("The TEMP environment variable is not defined.");
        }
        path = `${tempDir}\\OfficeAddins.log.txt`;
      }
      await registry.enableRuntimeLogging(path);
      return path;
  default:
    throw new Error(`Platform not supported: ${process.platform}.`);
  }
}

export async function getSourceBundleUrl(addinId: string): Promise<SourceBundleUrlComponents> {
  switch (process.platform) {
    case "win32":
      return registry.getSourceBundleUrl(addinId);
    default:
      throw new Error(`Platform not supported: ${process.platform}.`);
  }
}

export async function getEnabledDebuggingMethods(addinId: string): Promise<DebuggingMethod[]> {
  switch (process.platform) {
    case "win32":
      return registry.getEnabledDebuggingMethods(addinId);
    default:
      throw new Error(`Platform not supported: ${process.platform}.`);
  }
}

export async function getRuntimeLoggingPath(): Promise<string | undefined> {
  switch (process.platform) {
    case "win32":
      return registry.getRuntimeLoggingPath();
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

export async function setSourceBundleUrl(addinId: string, components: SourceBundleUrlComponents): Promise<void> {
  switch (process.platform) {
    case "win32":
      return registry.setSourceBundleUrl(addinId, components);
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
    .command("disable-runtime-logging")
    .description("Disables runtime logging.")
    .action(commands.disableRuntimeLogging);

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
    .command("enable-runtime-logging [path]")
    .description("Enable runtime logging.")
    .action(commands.enableRuntimeLogging);

  commander
    .command("get-source-bundle-url <manifestPath>")
    .description("Display the components of the url used to obtain the source bundle.")
    .action(commands.getSourceBundleUrl);

  commander
    .command("is-debugging-enabled [manifestPath]")
    .description("Display whether debugging is enabled for the add-in.")
    .action(commands.isDebuggingEnabled);

  commander
    .command("is-live-reload-enabled [manifestPath]")
    .description("Display whether live reload is enabled.")
    .action(commands.isDebuggingEnabled);

  commander
    .command("is-runtime-logging-enabled")
    .description("Display whether runtime logging is enabled.")
    .action(commands.isRuntimeLoggingEnabled);

  commander
    .command("set-source-bundle-url <manifestPath>")
    .description("Specify values for components of the url used to obtain the source bundle.")
    .option("-h,--host <host>", `The host name to use, or "" to use the default ('localhost').`)
    .option("-p,--port <port>", `The port number to use, or "" to use the default (8081).`)
    .option("--path <path>", `The path to use, or "" to use the default.`)
    .option("-e,--extension [extension]", `The extension to use ("" for no extension), or omit the value to use the default (".bundle").`)
    .action(commands.setSourceBundleUrl);

  commander.parse(process.argv);
}
