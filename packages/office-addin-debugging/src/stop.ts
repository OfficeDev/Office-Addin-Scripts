// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as fs from "fs";
import { parseNumber } from "office-addin-cli";
import { clearDevSettings} from "office-addin-dev-settings";
import { readManifestFile } from "office-addin-manifest";
import { startProcess, stopProcess } from "./process";
import { processIdFile } from "./start";
import { isNumber } from "util";

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

    const processId = readDevServerProcessId();
    if (processId) {
        stopProcess(processId);
        console.log(`Stopped dev server. Process id: ${processId}`);
        clearDevServerProcessId();
    }

    console.log("Debugging has been stopped.");
}

export function readDevServerProcessId(): number | undefined {
    let id;
    if (process.env.OfficeAddinDevServerProcessId) {
        id = parseNumber(process.env.OfficeAddinDevServerProcessId);
        console.log(`Process id read from env: ${id}`);
    } else if (fs.existsSync(processIdFile)) {
        const pid = fs.readFileSync(processIdFile);
        id = parseNumber(pid.toString(), `Invalid process id found in ${processIdFile}`);
        console.log(`Process id read from file: ${id}`);
    }
    return id;
}

export function clearDevServerProcessId() {
    console.log("Clearing dev server process id.");
    if (fs.existsSync(processIdFile)) {
        fs.unlinkSync(processIdFile);
    }
    delete process.env.OfficeAddinDevServerProcessId;
}