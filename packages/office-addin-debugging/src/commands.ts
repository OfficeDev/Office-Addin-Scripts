// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as commander from "commander";
import { logErrorMessage, parseNumber } from "office-addin-cli";
import * as devSettings from "office-addin-dev-settings";
import { OfficeApp, parseOfficeApp } from "office-addin-manifest";
import { AppType, parseAppType, parseDebuggingMethod, startDebugging } from "./start";
import { stopDebugging } from "./stop";

function getManifestPath(manifestPath: string, appTypeToDebug: string | undefined, prod: boolean): string {
    let manifestPathToDebug: string = manifestPath;
    try {
        if (appTypeToDebug === undefined) {
            throw new Error("Please specify the application type to debug.");
        }
        let platform: string = appTypeToDebug;
        if (appTypeToDebug === AppType.Desktop) {
            platform = process.platform;
        }
        if (manifestPath === "sdx" && platform !== AppType.Web) {
            if (process.env.npm_package_config_manifest_path) {
                manifestPathToDebug = process.env.npm_package_config_manifest_path;
                manifestPathToDebug = manifestPathToDebug.replace("${prodOrDev}", prod ? "prod" : "dev");
                manifestPathToDebug = manifestPathToDebug.replace("${platform}", platform);
            }
        }
    } catch (err) {
        logErrorMessage(`Unable to start debugging.\n${err}`);
    }
    return manifestPathToDebug;
}

function getSideloadCommand(sideload: string | undefined, manifestPath: string): string | undefined {
    let sideloadCommand: string | undefined = sideload;
    if (process.env.npm_package_config_sideload_command) {
        sideloadCommand = process.env.npm_package_config_sideload_command;
        sideloadCommand = sideloadCommand.replace("${manifestPath}", manifestPath);
    }
    return sideloadCommand;
}

function getUnloadCommand(unload: string | undefined, manifestPath: string): string | undefined {
    let unloadCommand: string | undefined = unload;
    if (process.env.npm_package_config_unload_command) {
        unloadCommand = process.env.npm_package_config_unload_command;
        unloadCommand = unloadCommand.replace("${manifestPath}", manifestPath);
    }
    return unloadCommand;
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

export async function start(manifestPath: string, appType: string | undefined, command: commander.Command) {
    try {
        const appTypeToDebug: AppType | undefined = parseAppType(appType || process.env.npm_package_config_app_type_to_debug || "desktop");
        const appToDebug: string | undefined = command.app || process.env.npm_package_config_app_to_debug;
        const app: OfficeApp | undefined = appToDebug ? parseOfficeApp(appToDebug) : undefined;
        const sourceBundleUrlComponents = new devSettings.SourceBundleUrlComponents(
            command.sourceBundleUrlHost, command.sourceBundleUrlPort,
            command.sourceBundleUrlPath, command.sourceBundleUrlExtension);
        const debuggingMethod = parseDebuggingMethod(command.debugMethod);
        const devServer: string | undefined = command.devServer || process.env.npm_package_scripts_dev_server;
        const devServerPort = parseDevServerPort(command.devServerPort || process.env.npm_package_config_dev_server_port);
        const enableDebugging: boolean = command.debug;
        const enableLiveReload: boolean = (command.liveReload === true);
        const manifestPathToDebug: string = getManifestPath(manifestPath, appTypeToDebug, command.prod);
        const packager: string | undefined = command.packager || process.env.npm_package_scripts_packager;
        const packagerHost: string | undefined = command.PackagerHost || process.env.npm_package_config_packager_host;
        const packagerPort: string | undefined = command.PackagerPort || process.env.npm_package_config_packager_port;

        if (appTypeToDebug === undefined) {
            throw new Error("Please specify the application type to debug.");
        }

        await startDebugging(manifestPath, appTypeToDebug, app, debuggingMethod, sourceBundleUrlComponents,
            devServer, devServerPort, packager, packagerHost, packagerPort, enableDebugging, enableLiveReload);
    } catch (err) {
        logErrorMessage(`Unable to start debugging.\n${err}`);
    }
}

export async function stop(manifestPath: string, appType: string | undefined, command: commander.Command) {
    try {
        await stopDebugging(manifestPath);
    } catch (err) {
        logErrorMessage(`Unable to stop debugging.\n${err}`);
    }
}
