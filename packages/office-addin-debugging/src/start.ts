// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as AdmZip from "adm-zip";
import * as fetch from "node-fetch";
import * as fspath from "path";
import * as devCerts from "office-addin-dev-certs";
import * as devSettings from "office-addin-dev-settings";
import * as os from "os";
import { DebuggingMethod, sideloadAddIn } from "office-addin-dev-settings";
import { OfficeApp, OfficeAddinManifest } from "office-addin-manifest";
import * as nodeDebugger from "office-addin-node-debugger";
import * as debugInfo from "./debugInfo";
import { getProcessIdsForPort } from "./port";
import { startDetachedProcess } from "./process";
import { usageDataObject } from "./defaults";
import { ExpectedError } from "office-addin-usage-data";

/* global process, console, setTimeout */

export enum AppType {
  Desktop = "desktop",
  Web = "web",
}

export enum Platform {
  Android = "android",
  Desktop = "desktop",
  iOS = "ios",
  MacOS = "macos",
  Win32 = "win32",
  Web = "web",
}

function defaultDebuggingMethod(): DebuggingMethod {
  return DebuggingMethod.Direct;
}

function delay(milliseconds: number): Promise<void> {
  return new Promise<void>((resolve) => {
    setTimeout(resolve, milliseconds);
  });
}

export async function isDevServerRunning(port: number): Promise<boolean> {
  // isPortInUse(port) will return false when webpack-dev-server is running.
  // it should be fixed, but for now, use getProcessIdsForPort(port)
  const processIds = await getProcessIdsForPort(port);
  const isRunning = processIds.length > 0;

  return isRunning;
}

export async function isPackagerRunning(statusUrl: string): Promise<boolean> {
  const statusRunningResponse = `packager-status:running`;

  try {
    const response = await fetch.default(statusUrl);
    console.log(`packager: ${response.status} ${response.statusText}`);
    const text = await response.text();
    console.log(`packager: ${text}`);
    return statusRunningResponse === text;
  } catch (err) {
    return false;
  }
}

export function parseDebuggingMethod(text: string): DebuggingMethod | undefined {
  switch (text) {
    case "direct":
      return DebuggingMethod.Direct;
    case "proxy":
      return DebuggingMethod.Proxy;
    default:
      return undefined;
  }
}

export function parsePlatform(text: string): Platform | undefined {
  if (text === AppType.Desktop) {
    text = process.platform;
  }

  switch (text) {
    case "android":
      return Platform.Android;
    case "darwin":
      return Platform.MacOS;
    case "ios":
      return Platform.iOS;
    case "macos":
      return Platform.MacOS;
    case "web":
      return Platform.Web;
    case "win32":
      return Platform.Win32;
    default:
      throw new ExpectedError(`The current platform is not supported: ${text}`);
  }
}

export async function runDevServer(commandLine: string, port?: number): Promise<void> {
  if (commandLine) {
    // if the dev server is running
    if (port !== undefined && (await isDevServerRunning(port))) {
      console.log(`The dev server is already running on port ${port}.`);
    } else {
      // On non-Windows platforms, prompt for installing the dev certs before starting the dev server.
      // This is a workaround for the fact that the detached process does not show a window on Mac,
      // therefore the user cannot enter the password when prompted.
      if (process.platform !== "win32") {
        if (!devCerts.verifyCertificates()) {
          await devCerts.ensureCertificatesAreInstalled();
        }
      }

      // start the dev server
      console.log(`Starting the dev server... (${commandLine})`);
      const devServerProcess = startDetachedProcess(commandLine);
      await debugInfo.saveDevServerProcessId(devServerProcess.pid);

      if (port !== undefined) {
        // wait until the dev server is running
        const isRunning: boolean = await waitUntilDevServerIsRunning(port);

        if (isRunning) {
          console.log(`The dev server is running on port ${port}. Process id: ${devServerProcess.pid}`);
        } else {
          throw new Error(`The dev server is not running on port ${port}.`);
        }
      }
    }
  }
}

export async function runNodeDebugger(host?: string, port?: string): Promise<void> {
  nodeDebugger.run(host, port);

  console.log("The node debugger is running.");
}

export async function runPackager(
  commandLine: string,
  host: string = "localhost",
  port: string = "8081"
): Promise<void> {
  if (commandLine) {
    const packagerUrl: string = `http://${host}:${port}`;
    const statusUrl: string = `${packagerUrl}/status`;

    // if the packager is running
    if (await isPackagerRunning(statusUrl)) {
      console.log(`The packager is already running. ${packagerUrl}`);
    } else {
      // start the packager
      console.log(`Starting the packager... (${commandLine})`);
      await startDetachedProcess(commandLine);

      // wait until the packager is running
      if (await waitUntilPackagerIsRunning(statusUrl)) {
        console.log(`The packager is running. ${packagerUrl}`);
      } else {
        throw new Error(`The packager is not running. ${packagerUrl}`);
      }
    }
  }
}

/**
 * `startDebugging` options.
 */
export interface StartDebuggingOptions {
  /**
   * The type of application to debug.
   */
  appType: AppType;

  /**
   * The Office application to debug.
   * If unspecified and there is more than one application in the manifest will prompt to specify the application.
   */
  app?: OfficeApp;

  /**
   * The method to use when debugging.
   */
  debuggingMethod?: DebuggingMethod;

  /**
   * Specify components of the source bundle url.
   */
  sourceBundleUrlComponents?: devSettings.SourceBundleUrlComponents;

  /**
   * If provided, starts the dev server.
   */
  devServerCommandLine?: string;

  /**
   *  If provided, port to verify that the dev server is running.
   */
  devServerPort?: number;

  /**
   *  If provided, starts the packager.
   */
  packagerCommandLine?: string;

  /**
   * Specifies the host name of the packager.
   */
  packagerHost?: string;

  /**
   * Specifies the port of the packager.
   */
  packagerPort?: string;

  /**
   * Enable debugging.
   * Starts with debugging if true or undefined.
   */
  enableDebugging?: boolean;

  /**
   * Enable live reload.
   */
  enableLiveReload?: boolean;

  /**
   * Enables launch of the Office host app and sideload of the add-in (if true or undefined).
   * Set to false to disable sideload.
   */
  enableSideload?: boolean;

  /**
   * Open Dev Tools.
   */
  openDevTools?: boolean;

  /**
   * If provided, the document to open for sideloading to web.
   */
  document?: string;
}

/**
 * Start debugging
 * @param options startDebugging options.
 */
export async function startDebugging(manifestPath: string, options: StartDebuggingOptions) {
  const {
    appType,
    app,
    debuggingMethod,
    sourceBundleUrlComponents,
    devServerCommandLine,
    devServerPort,
    packagerCommandLine,
    packagerHost,
    packagerPort,
    enableDebugging,
    enableLiveReload,
    enableSideload,
    openDevTools,
    document,
  } = {
    // Supplied Options
    ...options,

    // Defaults when variable is undefined.
    debuggingMethod: options.debuggingMethod || defaultDebuggingMethod(),
    enableDebugging: options.enableDebugging ?? true,
    enableSideload: options.enableSideload ?? true,
    enableLiveReload: options.enableLiveReload ?? true,
  };

  try {
    if (appType === undefined) {
      throw new ExpectedError("Please specify the application type to debug.");
    }

    const isWindowsPlatform = process.platform === "win32";
    const isDesktopAppType = appType === AppType.Desktop;
    const isProxyDebuggingMethod = debuggingMethod === DebuggingMethod.Proxy;

    // live reload can only be enabled for the desktop app type
    // when using proxy debugging and the packager
    const canEnableLiveReload: boolean = isDesktopAppType && isProxyDebuggingMethod && !!packagerCommandLine;
    // only use live reload if enabled and it can be enabled
    const useLiveReload = enableLiveReload && canEnableLiveReload;

    console.log(enableDebugging ? "Debugging is being started..." : "Starting without debugging...");
    console.log(`App type: ${appType}`);

    if (manifestPath.endsWith(".zip")) {
      manifestPath = await extractManifest(manifestPath);
    }
    const manifestInfo = await OfficeAddinManifest.readManifestFile(manifestPath);

    if (!manifestInfo.id) {
      throw new ExpectedError("Manifest does not contain the id for the Office Add-in.");
    }

    // enable loopback for Edge
    if (isWindowsPlatform && parseInt(os.release(), 10) === 10) {
      const name = isDesktopAppType ? "EdgeWebView" : "EdgeWebBrowser";
      await devSettings.ensureLoopbackIsEnabled(name);
    }

    // enable debugging
    if (isDesktopAppType && isWindowsPlatform) {
      await devSettings.enableDebugging(manifestInfo.id, enableDebugging, debuggingMethod, openDevTools);
      if (enableDebugging) {
        console.log(`Enabled debugging for add-in ${manifestInfo.id}.`);
      }
    }

    // enable live reload
    if (isDesktopAppType && isWindowsPlatform) {
      await devSettings.enableLiveReload(manifestInfo.id, useLiveReload);
      if (useLiveReload) {
        console.log(`Enabled live-reload for add-in ${manifestInfo.id}.`);
      }
    }

    // set source bundle url
    if (isDesktopAppType && isWindowsPlatform) {
      if (sourceBundleUrlComponents) {
        await devSettings.setSourceBundleUrl(manifestInfo.id, sourceBundleUrlComponents);
      }
    }

    // Run packager and dev server at the same time and wait for them to complete.
    let packagerPromise: Promise<void> | undefined;
    let devServerPromise: Promise<void> | undefined;

    if (packagerCommandLine && isProxyDebuggingMethod && isDesktopAppType) {
      packagerPromise = runPackager(packagerCommandLine, packagerHost, packagerPort);
    }

    if (devServerCommandLine) {
      devServerPromise = runDevServer(devServerCommandLine, devServerPort);
    }

    if (packagerPromise !== undefined) {
      try {
        await packagerPromise;
      } catch (err) {
        console.log(`Unable to start the packager. ${err}`);
      }
    }

    if (devServerPromise !== undefined) {
      try {
        await devServerPromise;
      } catch (err) {
        console.log(`Unable to start the dev server. ${err}`);
      }
    }

    if (enableDebugging && isProxyDebuggingMethod && isDesktopAppType) {
      try {
        await runNodeDebugger();
      } catch (err) {
        console.log(`Unable to start the node debugger. ${err}`);
      }
    }

    if (enableSideload) {
      try {
        console.log(`Sideloading the Office Add-in...`);
        await sideloadAddIn(manifestPath, app, true, appType, document);
      } catch (err) {
        throw new Error(`Unable to sideload the Office Add-in. \n${err}`);
      }
    }

    console.log(enableDebugging ? "Debugging started." : "Started.");

    usageDataObject.reportSuccess("startDebugging()", {
      app: app,
      document: document,
      appType: appType,
    });
  } catch (err: any) {
    usageDataObject.reportException("startDebugging()", err, {
      app: app,
      document: document,
      appType: appType,
    });
    throw err;
  }
}

export async function waitUntil(
  callback: () => Promise<boolean>,
  retryCount: number,
  retryDelay: number
): Promise<boolean> {
  let done: boolean = await callback();

  while (!done && retryCount) {
    --retryCount;
    await delay(retryDelay);
    done = await callback();
  }

  return done;
}

export async function waitUntilDevServerIsRunning(
  port: number,
  retryCount: number = 30,
  retryDelay: number = 1000
): Promise<boolean> {
  return waitUntil(async () => await isDevServerRunning(port), retryCount, retryDelay);
}

export async function waitUntilPackagerIsRunning(
  statusUrl: string,
  retryCount: number = 30,
  retryDelay: number = 1000
): Promise<boolean> {
  return waitUntil(async () => await isPackagerRunning(statusUrl), retryCount, retryDelay);
}

async function extractManifest(zipPath: string): Promise<string> {
  const targetPath: string = fspath.join(process.env.TEMP as string, "addinManifest");
  const zip = new AdmZip(zipPath); // reading archives
  zip.extractAllTo(targetPath, true); // overwrite
  return fspath.join(targetPath, "manifest.json");
}
