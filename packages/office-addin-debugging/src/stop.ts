// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as fs from "fs";
import { parseNumber } from "office-addin-cli";
import { clearDevSettings} from "office-addin-dev-settings";
import { readManifestFile } from "office-addin-manifest";
import { startProcess, stopProcess } from "./process";
import { processIdFile } from "./start";

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
 
    if (process.env.OfficeAddinDevServerProcessId) {
        const processId = parseNumber(process.env.OfficeAddinDevServerProcessId);

        if (processId) {
            stopProcess(processId);
            console.log(`Stopped dev server. Process id: ${processId}`);
        }

        delete process.env.OfficeAddinDevServerProcessId;
    }

    const processIdReadFromFile = parseNumber(readProcessIdFromFile());
    if (processIdReadFromFile) {
        stopProcess(processIdReadFromFile);
        console.log(`Stopped dev server. Process id: ${processIdReadFromFile}`);
        deleteProcessIdFile();
    }

    console.log("Debugging has been stopped.");
}

export function readProcessIdFromFile(): string | undefined {
    let id;
    if (fs.existsSync(processIdFile)) {
        id = fs.readFileSync(processIdFile);
        console.log(`Process id read from file: ${id}`);
    }
    return id ? id.toString() : undefined;
}

export function deleteProcessIdFile() {
    console.log(`Deleting process id file: ${processIdFile}`);
    fs.unlinkSync(processIdFile);
}