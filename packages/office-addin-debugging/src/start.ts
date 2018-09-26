import * as childProcess from "child_process";
import { ExecException } from "child_process";
import * as fetch from "node-fetch";
import * as devSettings from "office-addin-dev-settings";
import { DebuggingMethod } from "office-addin-dev-settings";
import * as manifest from "office-addin-manifest";
import * as nodeDebugger from "office-addin-node-debugger";

function defaultDebuggingMethod(): DebuggingMethod {
    switch (process.platform) {
        case "win32":
            return DebuggingMethod.Web; // will change to Direct later
        default:
            return DebuggingMethod.Web;
    }
}

function delay(milliseconds: number): Promise<void> {
    return new Promise<void>((resolve) => {
        setTimeout(resolve, milliseconds);
    });
}

export async function isDevServerRunning(url: string): Promise<boolean> {
    try {
        console.log(`Is dev server running? (${url})`);
        const response = await fetch.default(url);
        console.log(`devServer: ${response.status} ${response.statusText}`);
        return response.ok;
    } catch (err) {
        console.log(`Unable to connect to dev server.\n${err}`);
        return false;
    }
}

export async function isPackagerRunning(statusUrl: string): Promise<boolean> {
    const statusRunningResponse = `packager-status:running`;

    try {
        console.log(`Is packager running? (${statusUrl})`);
        const response = await fetch.default(statusUrl);
        console.log(`packager: ${response.status} ${response.statusText}`);
        const text = await response.text();
        console.log(`packager: ${text}`);
        return (statusRunningResponse === text);
    } catch (err) {
        return false;
    }
}

export function parseDebuggingMethod(text: string): DebuggingMethod {
    switch (text) {
        case "direct":
            return DebuggingMethod.Direct;
        default:
            return DebuggingMethod.Web;
    }
}

export async function runDevServer(commandLine: string, url?: string): Promise<void> {
    if (commandLine) {
        // if the dev server is running
        if (url && await isDevServerRunning(url)) {
            console.log(`The dev server is already running. ${url}`);
        } else {
            // start the dev server
            console.log(`Starting the dev server... (${commandLine})`);
            startProcess(commandLine);

            if (url) {
                // wait until the dev server is running
                if (await waitUntilDevServerIsRunning(url)) {
                    console.log(`The dev server is running. ${url}`);
                } else {
                    throw new Error("Unable to start the dev server.");
                }
            }
        }
    }
}

export async function runNodeDebugger(host?: string, port?: string): Promise<void> {
    nodeDebugger.run(host, port);
}

export async function runPackager(commandLine: string, host: string = "localhost", port: string = "8081"): Promise<void> {
    if (commandLine) {
        const packagerUrl: string = `http://${host}:${port}`;
        const statusUrl: string = `${packagerUrl}/status`;

        // if the packager is running
        if (await isPackagerRunning(statusUrl)) {
            console.log(`The packager is already running. ${packagerUrl}`);
        } else {
            // start the packager
            console.log(`Starting the packager... (${commandLine})`);
            startProcess(commandLine);

            // wait until the packager is running
            if (await waitUntilPackagerIsRunning(statusUrl)) {
                console.log(`The packager is running. ${packagerUrl}`);
            } else {
                throw new Error("Unable to start the packager.");
            }
        }
    }
}

export async function startDebugging(manifestPath: string,
    debuggingMethod: DebuggingMethod = defaultDebuggingMethod(),
    sourceBundleUrlComponents?: devSettings.SourceBundleUrlComponents,
    devServerCommandLine?: string, devServerUrl?: string,
    packagerCommandLine?: string, packagerHost?: string, packagerPort?: string,
    sideloadCommandLine: string = "npm run sideload") {

    let packagerPromise: Promise<void> | undefined;
    let devServerPromise: Promise<void> | undefined;

    console.log("Debugging is being started...");

    if (packagerCommandLine) {
        packagerPromise = runPackager(packagerCommandLine, packagerHost, packagerPort);
    }

    if (devServerCommandLine) {
        devServerPromise = runDevServer(devServerCommandLine, devServerUrl);
    }

    const manifestInfo = await manifest.readManifestFile(manifestPath);

    if (!manifestInfo.id) {
        throw new Error("Manifest does not contain the id for the Office Add-in.");
    }

    // enable debugging
    await devSettings.enableDebugging(manifestInfo.id, true, debuggingMethod);
    console.log(`Enabled debugging for add-in ${manifestInfo.id}. Debug method: ${debuggingMethod.toString()}`);

    // set source bundle url
    if (sourceBundleUrlComponents) {
        await devSettings.setSourceBundleUrl(manifestInfo.id, sourceBundleUrlComponents);
    }

    if (packagerPromise !== undefined) {
        try {
            await packagerPromise;
            console.log(`Started the packager.`);
        } catch (err) {
            console.log(`Unable to start the packager. ${err}`);
        }
    }

    if (devServerPromise !== undefined) {
        try {
            await devServerPromise;
            console.log(`Started the dev server.`);
        } catch (err) {
            console.log(`Unable to start the dev server. ${err}`);
        }
    }

    if (debuggingMethod === DebuggingMethod.Web) {
        try {
            await runNodeDebugger();
            console.log(`Started the node debugger.`);
        } catch (err) {
            console.log(`Unable to start the node debugger. ${err}`);
        }
    }

    try {
        console.log(`Sideloading the Office Add-in...`);
        await startProcess(sideloadCommandLine);
    } catch (err) {
        console.log(`Unable to sideload the Office Add-in. ${err}`);
    }

    console.log("Debugging started.");
}

async function startProcess(commandLine: string): Promise<void> {
    return new Promise<void>((resolve, reject) => {
        console.log(`Starting process: ${commandLine}`);
        childProcess.exec(commandLine, (error: ExecException | null, stdout: string, stderr: string) => {
            if (error) {
                reject(error);
            } else {
                resolve();
            }
        });
    });
}

export async function waitUntil(callback: (() => Promise<boolean>), retryCount: number, retryDelay: number): Promise<boolean> {
    let done: boolean = await callback();

    while (!done && retryCount) {
        --retryCount;
        delay(retryDelay);
        done = await callback();
    }

    return done;
}

export async function waitUntilDevServerIsRunning(url: string, retryCount: number = 10, retryDelay: number = 1000): Promise<boolean> {
    return waitUntil(() => {
        return isDevServerRunning(url);
    }, retryCount, retryDelay);
}

export async function waitUntilPackagerIsRunning(statusUrl: string, retryCount: number = 30, retryDelay: number = 1000): Promise<boolean> {
    return waitUntil(() => {
        return isPackagerRunning(statusUrl);
    }, retryCount, retryDelay);
}
