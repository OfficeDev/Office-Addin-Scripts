
// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import { readManifestFile } from "office-addin-manifest";
import * as devSettings from "./dev-settings";

export async function clear(manifestPath: string) {
    try {
        const manifest = await readManifestFile(manifestPath);

        if (manifest.id) {
            devSettings.clearDevSettings(manifest.id);
        } else {
            throw new Error(`The manifest file doesn't contain the id of the Office add-in.`);
        }
    } catch (err) {
        console.error(`Error: ${err}`);
    }
}

export async function disableDebugging(manifestPath: string) {
    try {
        const manifest = await readManifestFile(manifestPath);

        if (manifest.id) {
            devSettings.disableDebugging(manifest.id);
        } else {
            throw new Error(`The manifest file doesn't contain the id of the Office add-in.`);
        }
    } catch (err) {
        console.error(`Error: ${err}`);
    }
}

export async function enableDebugging(manifestPath: string) {
    try {
        const manifest = await readManifestFile(manifestPath);

        if (manifest.id) {
            devSettings.enableDebugging(manifest.id);
        } else {
            throw new Error(`The manifest file doesn't contain the id of the Office add-in.`);
        }
    } catch (err) {
        console.error(`Error: ${err}`);
    }
}
