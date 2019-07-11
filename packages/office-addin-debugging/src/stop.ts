// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as fs from "fs";
import { parseNumber } from "office-addin-cli";
import { clearDevSettings} from "office-addin-dev-settings";
import { readManifestFile } from "office-addin-manifest";
import { isNumber } from "util";
import * as debugInfo from "./debugInfo";
import { startProcess, stopProcess } from "./process";
import { processIdFilePath } from "./start";

export async function stopDebugging(manifestPath: string, unregisterCommandLine?: string) {
    console.log("Debugging is being stopped...");

    const isWindowsPlatform = (process.platform === "win32");
    const manifestInfo = await readManifestFile(manifestPath);

    if (!manifestInfo.id) {
        throw new Error("Manifest does not contain the id for the Office Add-in.");
    }

    // clear dev settings
    if (isWindowsPlatform) {
      await clearDevSettings(manifestInfo.id);
    }

    if (unregisterCommandLine) {
        // unregister
        try {
            await startProcess(unregisterCommandLine);
        } catch (err) {
            console.log(`Unable to unregister the Office Add-in. ${err}`);
        }
    }

    const processId = debugInfo.readDevServerProcessId();
    if (processId) {
        stopProcess(processId);
        console.log(`Stopped dev server. Process id: ${processId}`);
        debugInfo.clearDevServerProcessId();
    }

    console.log("Debugging has been stopped.");
}
