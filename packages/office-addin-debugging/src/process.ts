// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as childProcess from "child_process";
import { ExecException } from "child_process";

export async function startProcess(commandLine: string, verbose: boolean = false): Promise<void> {
    return new Promise<void>((resolve, reject) => {
        if (verbose) {
            console.log(`Starting: ${commandLine}`);
        }

        childProcess.exec(commandLine, (error: ExecException | null, stdout: string, stderr: string) => {
            if (error) {
                reject(error);
            } else {
                resolve();
            }
        });
    });
}

export function startDetachedProcess(commandLine: string, verbose: boolean = false): string {
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
    return subprocess.pid.toString();
}

export function stopProcess(): void {
    if (process.env.OfficeAddinDevServerProcessId) {
        try {
            if (process.platform === "win32") {
                childProcess.spawn("taskkill", ["/pid", process.env.OfficeAddinDevServerProcessId, "/f", "/t"]);
            } else {
                process.kill(Number(process.env.OfficeAddinDevServerProcessId));
            }
        } catch (err) {
            console.log(`Unable to kill process id ${process.env.OfficeAddinDevServerProcessId}: ${err}`);
        }
    }
}
