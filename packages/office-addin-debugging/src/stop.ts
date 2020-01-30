// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import { clearDevSettings, unregisterAddIn } from "office-addin-dev-settings";
import { readManifestFile } from "office-addin-manifest";
import * as debugInfo from "./debugInfo";
import { startProcess, stopProcess } from "./process";
import * as usageDataHelper from "./usagedata-helper";

export async function stopDebugging(manifestPath: string) {
    try {
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
    
        // unregister
        try {
            await unregisterAddIn(manifestPath);
        } catch (err) {
            console.log(`Unable to unregister the Office Add-in. ${err}`); //Do we want to throw or have separate telemetry here?
        }
    
        const processId = debugInfo.readDevServerProcessId();
        if (processId) {
            stopProcess(processId);
            console.log(`Stopped dev server. Process id: ${processId}`);
            debugInfo.clearDevServerProcessId();
        }
    
        console.log("Debugging has been stopped.");
        usageDataHelper.sendUsageDataSuccessEvent("stopDebugging");
    } catch(err) {
        usageDataHelper.sendUsageDataException("stopDebugging", err.message);
        throw err;
    }
}
