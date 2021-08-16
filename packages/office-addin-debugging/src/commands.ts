// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as commander from "commander";
import * as fs from "fs";
import { logErrorMessage, parseNumber, getPackageJsonScript } from "office-addin-cli";
import * as devSettings from "office-addin-dev-settings";
import { OfficeApp, parseOfficeApp } from "office-addin-manifest";
import { AppType, parseDebuggingMethod, parsePlatform, Platform, startDebugging } from "./start";
import { stopDebugging } from "./stop";
import { usageDataObject } from "./defaults";
import { ExpectedError } from "office-addin-usage-data";

/* global process */

function determineManifestPath(platform: Platform, dev: boolean): string {
  let manifestPath = process.env.npm_package_config_manifest_location || "";
  manifestPath = manifestPath.replace("${flavor}", dev ? "dev" : "prod").replace("${platform}", platform);

  if (!manifestPath) {
    throw new ExpectedError(`The manifest path was not provided.`);
  }
  if (!fs.existsSync(manifestPath)) {
    throw new ExpectedError(`The manifest path does not exist: ${manifestPath}.`);
  }

  return manifestPath;
}

function parseDevServerPort(optionValue: any): number | undefined {
  const devServerPort = parseNumber(optionValue, "--dev-server-port should specify a number.");

  if (devServerPort !== undefined) {
    if (!Number.isInteger(devServerPort)) {
      throw new ExpectedError("--dev-server-port should be an integer.");
    }
    if (devServerPort < 0 || devServerPort > 65535) {
      throw new ExpectedError("--dev-server-port should be between 0 and 65535.");
    }
  }

  return devServerPort;
}

export async function start(manifestPath: string, platform: string | undefined, command: commander.Command) {
  try {
    const appPlatformToDebug: Platform | undefined = parsePlatform(
      platform || process.env.npm_package_config_app_platform_to_debug || Platform.Win32
    );
    const appTypeToDebug: AppType | undefined = devSettings.parseAppType(
      appPlatformToDebug || process.env.npm_package_config_app_type_to_debug || AppType.Desktop
    );
    const appToDebug: string | undefined = command.app || process.env.npm_package_config_app_to_debug;
    const app: OfficeApp | undefined = appToDebug ? parseOfficeApp(appToDebug) : undefined;
    const dev: boolean = command.prod ? false : true;
    const debuggingMethod = parseDebuggingMethod(command.debugMethod);
    const devServer: string | undefined = command.devServer || (await getPackageJsonScript("dev-server"));
    const devServerPort = parseDevServerPort(command.devServerPort || process.env.npm_package_config_dev_server_port);
    const document: string | undefined = command.document || process.env.npm_package_config_document;
    const enableDebugging: boolean = command.debug;
    const enableLiveReload: boolean = command.liveReload === true;
    const enableSideload: boolean = command.sideload !== false; // enable if true or undefined; only disable if false
    const openDevTools: boolean = command.devTools === true;
    const packager: string | undefined = command.packager || (await getPackageJsonScript("packager"));
    const packagerHost: string | undefined = command.PackagerHost || process.env.npm_package_config_packager_host;
    const packagerPort: string | undefined = command.PackagerPort || process.env.npm_package_config_packager_port;
    const sourceBundleUrlComponents = new devSettings.SourceBundleUrlComponents(
      command.sourceBundleUrlHost,
      command.sourceBundleUrlPort,
      command.sourceBundleUrlPath,
      command.sourceBundleUrlExtension
    );

    if (appPlatformToDebug === undefined) {
      throw new ExpectedError("Please specify the platform to debug.");
    }

    if (appTypeToDebug === undefined) {
      throw new ExpectedError("Please specify the application type to debug.");
    }

    if (appPlatformToDebug === Platform.Android || appPlatformToDebug === Platform.iOS) {
      throw new ExpectedError(`Platform type ${appPlatformToDebug} not currently supported for debugging`);
    }

    if (!manifestPath) {
      manifestPath = determineManifestPath(appPlatformToDebug, dev);
    }

    await startDebugging(manifestPath, {
      appType: appTypeToDebug,
      app,
      debuggingMethod,
      sourceBundleUrlComponents,
      devServerCommandLine: devServer,
      devServerPort,
      packagerCommandLine: packager,
      packagerHost,
      packagerPort,
      enableDebugging,
      enableLiveReload,
      enableSideload,
      openDevTools,
      document,
    });

    usageDataObject.reportSuccess("start");
  } catch (err) {
    usageDataObject.reportException("start", err);
    logErrorMessage(`Unable to start debugging.\n${err}`);
  }
}

export async function stop(manifestPath: string, platform: string | undefined, command: commander.Command) {
  try {
    const appPlatformToDebug: Platform | undefined = parsePlatform(
      platform || process.env.npm_package_config_app_plaform_to_debug || Platform.Win32
    );
    const dev: boolean = command.prod ? false : true;

    if (manifestPath === "" && appPlatformToDebug !== undefined) {
      manifestPath = determineManifestPath(appPlatformToDebug, dev);
    }

    await stopDebugging(manifestPath);
    usageDataObject.reportSuccess("stop");
  } catch (err) {
    usageDataObject.reportException("stop", err);
    logErrorMessage(`Unable to stop debugging.\n${err}`);
  }
}
