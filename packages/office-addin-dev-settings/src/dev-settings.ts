// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as fs from "fs";
import * as fspath from "path";
import * as registry from "./dev-settings-registry";

const defaultRuntimeLogFileName = "OfficeAddins.log.txt";

export enum DebuggingMethod {
  Direct,
  Proxy,
  /** @deprecated use Proxy */
  Web = Proxy,
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

export async function enableDebugging(addinId: string, enable: boolean = true, method: DebuggingMethod = DebuggingMethod.Direct): Promise<void> {
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
