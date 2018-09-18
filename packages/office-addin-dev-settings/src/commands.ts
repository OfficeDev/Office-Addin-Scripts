
// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as commander from "commander";
import { ManifestInfo, readManifestFile } from "office-addin-manifest";
import * as devSettings from "./dev-settings";

export async function clear(manifestPath: string) {
    try {
        const manifest = await readManifestFile(manifestPath);

        validateManifestId(manifest);

        devSettings.clearDevSettings(manifest.id!);
    } catch (err) {
        logErrorMessage(err);
    }
}

export async function disableDebugging(manifestPath: string) {
    try {
        const manifest = await readManifestFile(manifestPath);

        validateManifestId(manifest);

        devSettings.disableDebugging(manifest.id!);
    } catch (err) {
        logErrorMessage(err);
    }
}

export async function disableLiveReload(manifestPath: string) {
    try {
        const manifest = await readManifestFile(manifestPath);

        validateManifestId(manifest);

        devSettings.disableLiveReload(manifest.id!);
    } catch (err) {
        logErrorMessage(err);
    }
}

export async function enableDebugging(manifestPath: string, command: commander.Command) {
    try {
        const manifest = await readManifestFile(manifestPath);

        validateManifestId(manifest);

        devSettings.enableDebugging(manifest.id!, true, toDebuggingMethod(command.debugMethod));
    } catch (err) {
        logErrorMessage(err);
    }
}

export async function enableLiveReload(manifestPath: string) {
    try {
        const manifest = await readManifestFile(manifestPath);

        validateManifestId(manifest);

        devSettings.enableLiveReload(manifest.id!);
    } catch (err) {
        logErrorMessage(err);
    }
}

function logErrorMessage(err: any) {
    console.error(`Error: ${err instanceof Error ? err.message : err}`);
}

function toDebuggingMethod(text?: string): devSettings.DebuggingMethod {
    switch (text) {
        case "direct":
            return devSettings.DebuggingMethod.Direct;
        case "web":
            return devSettings.DebuggingMethod.Web;
        case "":
        case null:
        case undefined:
            // preferred debug method
            return devSettings.DebuggingMethod.Web;
        default:
            throw new Error(`Please provide a valid debug method instead of '${text}'.`);
    }
}

function validateManifestId(manifest: ManifestInfo) {
    if (!manifest.id) {
        throw new Error(`The manifest file doesn't contain the id of the Office add-in.`);
    }
}
