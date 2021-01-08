// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as fs from "fs";
import { readManifestFile } from "office-addin-manifest";
import * as fspath from "path";
import * as devSettingsMac from "./dev-settings-mac";
import * as devSettingsWindows from "./dev-settings-windows";
import { platform } from "os";

const defaultRuntimeLogFileName = "OfficeAddins.log.txt";

export enum DebuggingMethod {
  Direct,
  Proxy,
  /** @deprecated use Proxy */
  Web = Proxy,
}

export enum WebViewType {
  Default = "Default",
  IE = "IE",
  Edge = "Edge",
  EdgeChromium = "Edge Chromium",
}

export class RegisteredAddin {
  public id: string;
  public manifestPath: string;

  constructor(id: string, manifestPath: string) {
    this.id = id;
    this.manifestPath = manifestPath;
  }
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

    return `http://${host}${host && port ? ":" : ""}${port}/${path}${extension}`;
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
      return devSettingsWindows.clearDevSettings(addinId);
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
      return devSettingsWindows.disableRuntimeLogging();
  default:
    throw new Error(`Platform not supported: ${process.platform}.`);
  }
}

export async function enableDebugging(addinId: string, enable: boolean = true, method: DebuggingMethod = DebuggingMethod.Direct, openDevTools = false): Promise<void> {
  switch (process.platform) {
    case "win32":
      return devSettingsWindows.enableDebugging(addinId, enable, method, openDevTools);
    default:
      throw new Error(`Platform not supported: ${process.platform}.`);
  }
}

export async function enableLiveReload(addinId: string, enable: boolean = true): Promise<void> {
  switch (process.platform) {
    case "win32":
      return devSettingsWindows.enableLiveReload(addinId, enable);
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
        path = fspath.normalize(`${tempDir}/${defaultRuntimeLogFileName}`);
      }

      const pathExists: boolean = fs.existsSync(path);
      if (pathExists) {
        const stat = fs.statSync(path);
        if (stat.isDirectory()) {
          throw new Error(`You need to specify the path to a file. This is a directory: "${path}".`);
        }
      }
      try {
        const file = fs.openSync(path, "a+");
        fs.closeSync(file);
      } catch (err) {
        throw new Error(pathExists
          ? `You need to specify the path to a writable file. Unable to write to: "${path}".`
          : `You need to specify the path where the file can be written. Unable to write to: "${path}".`);
      }

      await devSettingsWindows.enableRuntimeLogging(path);
      return path;
  default:
    throw new Error(`Platform not supported: ${process.platform}.`);
  }
}

/**
 * Returns the manifest paths for the add-ins that are registered
 */
export async function getRegisterAddIns(): Promise<RegisteredAddin[]>  {
  switch (process.platform) {
    case "darwin":
      return devSettingsMac.getRegisteredAddIns();
    case "win32":
      return devSettingsWindows.getRegisteredAddIns();
    default:
      throw new Error(`Platform not supported: ${process.platform}.`);
  }
}

export async function getEnabledDebuggingMethods(addinId: string): Promise<DebuggingMethod[]> {
  switch (process.platform) {
    case "win32":
      return devSettingsWindows.getEnabledDebuggingMethods(addinId);
    default:
      throw new Error(`Platform not supported: ${process.platform}.`);
  }
}

export async function getOpenDevTools(addinId: string): Promise<boolean> {
  switch (process.platform) {
    case "win32":
      return devSettingsWindows.getOpenDevTools(addinId);
    default:
      throw new Error(`Platform not supported: ${process.platform}.`);
  }
}

export async function getRuntimeLoggingPath(): Promise<string | undefined> {
  switch (process.platform) {
    case "win32":
      return devSettingsWindows.getRuntimeLoggingPath();
    default:
      throw new Error(`Platform not supported: ${process.platform}.`);
  }
}

export async function getSourceBundleUrl(addinId: string): Promise<SourceBundleUrlComponents> {
  switch (process.platform) {
    case "win32":
      return devSettingsWindows.getSourceBundleUrl(addinId);
    default:
      throw new Error(`Platform not supported: ${process.platform}.`);
  }
}

export async function getWebView(addinId: string): Promise<WebViewType | undefined> {
  switch (process.platform) {
    case "win32":
      return devSettingsWindows.getWebView(addinId);
    default:
      throw new Error(`Platform not supported: ${process.platform}.`);
  }
}

export async function isDebuggingEnabled(addinId: string): Promise<boolean> {
  switch (process.platform) {
    case "win32":
      return devSettingsWindows.isDebuggingEnabled(addinId);
    default:
      throw new Error(`Platform not supported: ${process.platform}.`);
  }
}

export async function isLiveReloadEnabled(addinId: string): Promise<boolean> {
  switch (process.platform) {
    case "win32":
      return devSettingsWindows.isLiveReloadEnabled(addinId);
    default:
      throw new Error(`Platform not supported: ${process.platform}.`);
  }
}

export async function registerAddIn(manifestPath: string): Promise<void> {
  switch (process.platform) {
    case "win32":
      const manifest = await readManifestFile(manifestPath);
      const realManifestPath = fs.realpathSync(manifestPath);
      return devSettingsWindows.registerAddIn(manifest.id || "", realManifestPath);
    case "darwin":
      return devSettingsMac.registerAddIn(manifestPath);
    default:
      throw new Error(`Platform not supported: ${process.platform}.`);
  }
}

export async function setSourceBundleUrl(addinId: string, components: SourceBundleUrlComponents): Promise<void> {
  switch (process.platform) {
    case "win32":
      return devSettingsWindows.setSourceBundleUrl(addinId, components);
    default:
      throw new Error(`Platform not supported: ${process.platform}.`);
  }
}

export async function setWebView(addinId: string, webViewType: WebViewType | undefined) {
  switch (process.platform) {
    case "win32":
      return devSettingsWindows.setWebView(addinId, webViewType);;
    default:
      throw new Error(`Platform not supported: ${process.platform}.`);
  }
}

export async function unregisterAddIn(manifestPath: string): Promise<void> {
  switch (process.platform) {
    case "darwin":
      return devSettingsMac.unregisterAddIn(manifestPath);
    case "win32":
      const manifest = await readManifestFile(manifestPath);
      const realManifestPath = fs.realpathSync(manifestPath);
      return devSettingsWindows.unregisterAddIn(manifest.id || "", realManifestPath);
    default:
      throw new Error(`Platform not supported: ${process.platform}.`);
  }
}

export async function unregisterAllAddIns(): Promise<void> {
  switch (process.platform) {
    case "darwin":
      return devSettingsMac.unregisterAllAddIns();
    case "win32":
      return devSettingsWindows.unregisterAllAddIns();
    default:
      throw new Error(`Platform not supported: ${process.platform}.`);
  }
}
