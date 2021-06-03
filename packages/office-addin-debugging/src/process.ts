// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import { ChildProcess } from "child_process";
import * as devSettings from "office-addin-dev-settings";

export async function startProcess(
  commandLine: string,
  verbose: boolean = false
): Promise<void> {
  await devSettings.startProcess(commandLine, verbose);
}

export function startDetachedProcess(
  commandLine: string,
  verbose: boolean = false
): ChildProcess {
  return devSettings.startDetachedProcess(commandLine, verbose);
}

export function stopProcess(processId: number): void {
  devSettings.stopProcess(processId);
}
