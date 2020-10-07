// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as commander from "commander";
import * as fs from "fs";
import { logErrorMessage, parseNumber } from "office-addin-cli";
import * as devSettings from "office-addin-dev-settings";
import { OfficeApp, parseOfficeApp } from "office-addin-manifest";
import { AppType, parseAppType, parseDebuggingMethod, parsePlatform, Platform, startDebugging } from "./start";
import { stopDebugging } from "./stop";

function determineManifestPath(platform: Platform, dev: boolean): string {
    let manifestPath = process.env.npm_package_config_manifest_location || "";
    manifestPath = manifestPath.replace("${flavor}", dev ? "dev" : "prod").replace("${platform}", platform);

    if (!manifestPath) {
      throw new Error(`The manifest path was not provided.`);
    }
    if (!fs.existsSync(manifestPath)) {
      throw new Error(`The manifest path does not exist: ${manifestPath}.`);
    }

    return manifestPath;
}

function parseDevServerPort(optionValue: any): number | undefined {
    const devServerPort = parseNumber(optionValue, "--dev-server-port should specify a number.");

    if (devServerPort !== undefined) {
        if (!Number.isInteger(devServerPort)) {
            throw new Error("--dev-server-port should be an integer.");
        }
        if ((devServerPort < 0) || (devServerPort > 65535)) {
            throw new Error("--dev-server-port should be between 0 and 65535.");
        }
    }

    return devServerPort;
}

export async function start(manifestPath: string, platform: string | undefined, command: commander.Command) {
    try {
        const appPlatformToDebug: Platform | undefined = parsePlatform(platform || process.env.npm_package_config_app_platform_to_debug || Platform.Win32);
        const appTypeToDebug: AppType | undefined = parseAppType(appPlatformToDebug || process.env.npm_package_config_app_type_to_debug || AppType.Desktop);
        const appToDebug: string | undefined = command.app || process.env.npm_package_config_app_to_debug;
        const app: OfficeApp | undefined = appToDebug ? parseOfficeApp(appToDebug) : undefined;
        const dev: boolean = command.prod ? false : true;
        const debuggingMethod = parseDebuggingMethod(command.debugMethod);
        const devServer: string | undefined = command.devServer || process.env.npm_package_scripts_dev_server;
        const devServerPort = parseDevServerPort(command.devServerPort || process.env.npm_package_config_dev_server_port);
        const enableDebugging: boolean = command.debug;
        const enableLiveReload: boolean = (command.liveReload === true);
        const openDevTools: boolean = (command.devTools === true);
        const packager: string | undefined = command.packager || process.env.npm_package_scripts_packager;
        const packagerHost: string | undefined = command.PackagerHost || process.env.npm_package_config_packager_host;
        const packagerPort: string | undefined = command.PackagerPort || process.env.npm_package_config_packager_port;
        const skipLoopbackPrompt: boolean  | undefined = command.skipLoopbackPrompt;
        const sourceBundleUrlComponents = new devSettings.SourceBundleUrlComponents(
            command.sourceBundleUrlHost, command.sourceBundleUrlPort,
            command.sourceBundleUrlPath, command.sourceBundleUrlExtension);

        if (appPlatformToDebug === undefined) {
            throw new Error("Please specify the platform to debug.");
        }

        if (appTypeToDebug === undefined) {
            throw new Error("Please specify the application type to debug.");
        }

        if (appPlatformToDebug === Platform.Android || appPlatformToDebug === Platform.iOS) {
            throw new Error(`Platform type ${appPlatformToDebug} not currently supported for debugging`);
        }

        if (!manifestPath) {
            manifestPath = determineManifestPath(appPlatformToDebug, dev);
        }

        await startDebugging(manifestPath, appTypeToDebug, app, debuggingMethod, sourceBundleUrlComponents,
            devServer, devServerPort, packager, packagerHost, packagerPort, enableDebugging, enableLiveReload, openDevTools, skipLoopbackPrompt);
    } catch (err) {
        logErrorMessage(`Unable to start debugging.\n${err}`);
    }
}

export async function stop(manifestPath: string, platform: string | undefined, command: commander.Command) {
    try {
        const appPlatformToDebug: Platform | undefined = parsePlatform(platform || process.env.npm_package_config_app_plaform_to_debug || Platform.Win32);
        const dev: boolean = command.prod ? false : true;

        if (manifestPath === "" && appPlatformToDebug !== undefined) {
            manifestPath = determineManifestPath(appPlatformToDebug, dev);
        }

        await stopDebugging(manifestPath);
    } catch (err) {
        logErrorMessage(`Unable to stop debugging.\n${err}`);
    }
}
