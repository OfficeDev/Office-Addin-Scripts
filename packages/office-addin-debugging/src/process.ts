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

export function startDetachedProcess(commandLine: string, verbose: boolean = false): void {
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
}
