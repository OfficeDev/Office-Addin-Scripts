// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as childProcess from "child_process";
import { ChildProcess, ExecException } from "child_process";

/* global process, console */

export async function startProcess(commandLine: string, verbose: boolean = false): Promise<void> {
  return new Promise<void>((resolve, reject) => {
    if (verbose) {
      console.log(`Starting: ${commandLine}`);
    }

    childProcess.exec(
      commandLine,
      (
        error: ExecException | null,
        stdout: string /* eslint-disable-line no-unused-vars */,
        stderr: string /* eslint-disable-line no-unused-vars */
      ) => {
        if (error) {
          reject(error);
        } else {
          resolve();
        }
      }
    );
  });
}

export function startDetachedProcess(commandLine: string, verbose: boolean = false): ChildProcess {
  if (verbose) {
    console.log(`Starting: ${commandLine}`);
  }

  const subprocess = childProcess.spawn(commandLine, [], {
    detached: true,
    shell: true,
    stdio: "ignore",
    windowsHide: false,
  });

  subprocess.on("error", (err) => {
    console.log(`Unable to run command: ${commandLine}.\n${err}`);
  });

  subprocess.unref();
  return subprocess;
}

export function stopProcess(processId: number): void {
  if (processId) {
    try {
      if (process.platform === "win32") {
        childProcess.spawn("taskkill", ["/pid", `${processId}`, "/f", "/t"]);
      } else {
        process.kill(processId);
      }
    } catch (err) {
      console.log(`Unable to kill process id ${processId}: ${err}`);
    }
  }
}
