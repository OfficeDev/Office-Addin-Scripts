// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as fs from "fs";
import { parseNumber } from "office-addin-cli";
import * as os from "os";
import * as path from "path";

const processIdFile = "office-addin-debugging.json";
const processIdFilePath = path.join(os.tmpdir(), processIdFile);

export interface IDebugInfo {
    devServer: IDevServerInfo;
}

export interface IDevServerInfo {
    processId: number;
}

export function getDebuggingInfoPath(): string {
    return processIdFilePath;
}

/**
 * Creates the DebugInfo object
 */
function createDebugInfo(): IDebugInfo {
    const devServer: IDevServerInfo = {
        processId : 0,
    };
    const debugInfo: IDebugInfo = {
        devServer,
    };
    return debugInfo;
}

/**
 * Set the process id of the dev server in the debuginfo object
 * @param debugInfo DebugInfo object
 * @param id dev server process id
 */
function setDevServerProcessId(debugInfo: IDebugInfo, id: number) {
    debugInfo.devServer.processId = id;
}

/**
 * Loads the DebugInfo object into JSON
 * @param debugInfo DebugInfo object
 */
function loadDebugInfo(debugInfo: IDebugInfo): string {
    return JSON.stringify(debugInfo, null, 4 );
}

/**
 * Read the DebugInfo object
 * @param pathToFile - Path to the json file containing debug info
 */
function readDebugInfo(pathToFile: string): any {
    const json = fs.readFileSync(pathToFile);
    return JSON.parse(json.toString());
}

/**
 * Saves the process id of the dev server
 * @param id process id
 */
export async function saveDevServerProcessId(id: number): Promise<void> {
    console.log(`Writing process id: ${id} to file: ${processIdFilePath}`);
    process.env.OfficeAddinDevServerProcessId = id.toString();
    const debugInfo = createDebugInfo();
    setDevServerProcessId(debugInfo, id);
    fs.writeFileSync(processIdFilePath, loadDebugInfo(debugInfo));
}

/**
 * Read the process id from either the environment variable or
 * from the json file containing the debug info
 */
export function readDevServerProcessId(): number | undefined {
    let id;
    if (process.env.OfficeAddinDevServerProcessId) {
        id = parseNumber(process.env.OfficeAddinDevServerProcessId);
        console.log(`Process id read from env: ${id}`);
    } else if (fs.existsSync(processIdFilePath)) {
        const devServerProperties = readDebugInfo(processIdFilePath);
        const pid = devServerProperties.devServer.processId;
        id = parseNumber(pid.toString(), `Invalid process id found in ${processIdFilePath}`);
        console.log(`Process id read from file: ${id}`);
    }
    return id;
}

/**
 * Deletes the environment variable containing process id
 * and deletes the debug info json file if it exists
 */
export function clearDevServerProcessId() {
    console.log("Clearing dev server process id.");
    if (fs.existsSync(processIdFilePath)) {
        fs.unlinkSync(processIdFilePath);
    }
    delete process.env.OfficeAddinDevServerProcessId;
}
