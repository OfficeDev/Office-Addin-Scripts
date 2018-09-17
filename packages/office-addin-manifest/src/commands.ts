// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import { readManifestFile } from "./manifest";

export async function info(path: string) {
    try {
        const manifest = await readManifestFile(path);

        console.log(`Manifest: ${path}`);
        console.log(`  Id: ${manifest.id || ""}`);
        console.log(`  Name: ${manifest.displayName || ""}`);
        console.log(`  Provider: ${manifest.providerName || ""}`);
        console.log(`  Type: ${manifest.officeAppType || ""}`);
        console.log(`  Version: ${manifest.version || ""}`);
        console.log(`  Default Locale: ${manifest.defaultLocale || ""}`);
        console.log(`  Description: ${manifest.description || ""}`);
    } catch (err) {
        console.error(`Error: ${err}`);
    }
}
