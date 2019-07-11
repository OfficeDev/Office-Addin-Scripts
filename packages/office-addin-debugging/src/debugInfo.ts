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
 * @param id - process id of the dev server
 */
function createDebugInfoObject(id: number): IDebugInfo {
    const devServer: IDevServerInfo = {
        processId : id,
    };
    const debugInfo: IDebugInfo = {
        devServer,
    };
    return debugInfo;
}

/**
 * Read the DebugInfo object
 * @param pathToFile - Path to the json file containing debug info
 */
function readDebugInfoObject(pathToFile: string): any {
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
    const json = JSON.stringify(createDebugInfoObject(id), null, 4 );
    fs.writeFileSync(processIdFilePath, json);
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
        const devServerProperties = readDebugInfoObject(processIdFilePath);
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
