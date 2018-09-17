
// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import { ManifestInfo, readManifestFile } from "office-addin-manifest";
import * as devSettings from "./dev-settings";

export async function clear(manifestPath: string) {
    try {
        const manifest = await readManifestFile(manifestPath);

        validateManifestId(manifest);

        devSettings.clearDevSettings(manifest.id!);
    } catch (err) {
        console.error(`Error: ${err}`);
    }
}

export async function disableDebugging(manifestPath: string) {
    try {
        const manifest = await readManifestFile(manifestPath);

        validateManifestId(manifest);

        devSettings.disableDebugging(manifest.id!);
    } catch (err) {
        console.error(`Error: ${err}`);
    }
}

export async function enableDebugging(manifestPath: string) {
    try {
        const manifest = await readManifestFile(manifestPath);

        validateManifestId(manifest);

        devSettings.enableDebugging(manifest.id!);
    } catch (err) {
        console.error(`Error: ${err}`);
    }
}

function validateManifestId(manifest: ManifestInfo) {
    if (!manifest.id) {
        throw new Error(`The manifest file doesn't contain the id of the Office add-in.`);
    }
}
