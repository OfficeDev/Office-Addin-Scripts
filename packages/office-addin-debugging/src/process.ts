import * as childProcess from "child_process";
import { ExecException } from "child_process";
import { getProcessIdsForPort } from "./port";

export async function startProcess(commandLine: string, verbose: boolean = false): Promise<void> {
    return new Promise<void>((resolve, reject) => {
        if (verbose) {
            console.log(`Starting process: ${commandLine}`);
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
