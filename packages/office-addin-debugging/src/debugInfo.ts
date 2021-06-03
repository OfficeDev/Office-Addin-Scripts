// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as fs from "fs";
import { parseNumber } from "office-addin-cli";
import * as os from "os";
import * as path from "path";

/* global process */

const processIdFile = "office-addin-debugging.json";
const processIdFilePath = path.join(os.tmpdir(), processIdFile);

export interface IDebuggingInfo {
  devServer: IDevServerInfo;
}

export interface IDevServerInfo {
  processId: number | undefined;
}

export function getDebuggingInfoPath(): string {
  return processIdFilePath;
}

/**
 * Creates the DebuggingInfo object
 */
function createDebuggingInfo(): IDebuggingInfo {
  const devServer: IDevServerInfo = {
    processId: undefined,
  };
  const debuggingInfo: IDebuggingInfo = {
    devServer,
  };
  return debuggingInfo;
}

/**
 * Set the process id of the dev server in the Debugging object
 * @param debuggingInfo DebuggingInfo object
 * @param id dev server process id
 */
function setDevServerProcessId(debuggingInfo: IDebuggingInfo, id: number) {
  debuggingInfo.devServer.processId = id;
}

/**
 * Write the DebuggingInfo to the json file
 * @param debuggingInfo DebuggingInfo object
 * @param pathToFile Path to the json file containing debug info
 */
function writeDebuggingInfo(debuggingInfo: IDebuggingInfo, pathToFile: string) {
  const debuggingInfoJSON = JSON.stringify(debuggingInfo, null, 4);
  fs.writeFileSync(pathToFile, debuggingInfoJSON);
}

/**
 * Read the DebuggingInfo object
 * @param pathToFile Path to the json file containing debug info
 */
function readDebuggingInfo(pathToFile: string): any {
  const json = fs.readFileSync(pathToFile);
  return JSON.parse(json.toString());
}

/**
 * Saves the process id of the dev server
 * @param id process id
 */
export async function saveDevServerProcessId(id: number): Promise<void> {
  process.env.OfficeAddinDevServerProcessId = id.toString();
  const debuggingInfo = createDebuggingInfo();
  setDevServerProcessId(debuggingInfo, id);
  writeDebuggingInfo(debuggingInfo, processIdFilePath);
}

/**
 * Read the process id from either the environment variable or
 * from the json file containing the debug info
 */
export function readDevServerProcessId(): number | undefined {
  let id;
  if (process.env.OfficeAddinDevServerProcessId) {
    id = parseNumber(process.env.OfficeAddinDevServerProcessId);
  } else if (fs.existsSync(processIdFilePath)) {
    const devServerProperties = readDebuggingInfo(processIdFilePath);
    if (
      devServerProperties.devServer &&
      devServerProperties.devServer.processId
    ) {
      const pid = devServerProperties.devServer.processId;
      id = parseNumber(
        pid.toString(),
        `Invalid process id found in ${processIdFilePath}`
      );
    }
  }
  return id;
}

/**
 * Deletes the environment variable containing process id
 * and deletes the debug info json file if it exists
 */
export function clearDevServerProcessId() {
  if (fs.existsSync(processIdFilePath)) {
    fs.unlinkSync(processIdFilePath);
  }
  delete process.env.OfficeAddinDevServerProcessId;
}
