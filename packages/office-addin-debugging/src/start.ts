import * as fetch from "node-fetch";
import * as devSettings from "office-addin-dev-settings";
import { DebuggingMethod } from "office-addin-dev-settings";
import * as manifest from "office-addin-manifest";
import * as nodeDebugger from "office-addin-node-debugger";
import { isPortInUse } from "./port";
import { startProcess } from "./process";

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

export async function isDevServerRunning(port: number): Promise<boolean> {
    return isPortInUse(port);
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

export async function runDevServer(commandLine: string, port?: number): Promise<void> {
    if (commandLine) {
        // if the dev server is running
        if ((port !== undefined) && await isDevServerRunning(port)) {
            console.log("The dev server is already running.");
        } else {
            // start the dev server
            console.log(`Starting the dev server... (${commandLine})`);
            startProcess(commandLine);

            if (port !== undefined) {
                // wait until the dev server is running
                if (await waitUntilDevServerIsRunning(port)) {
                    console.log(`The dev server is running.`);
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

/**
 * Start debugging
 * @param manifestPath The path to the manifest file.
 * @param debuggingMethod The method to use when debugging.
 * @param sourceBundleUrlComponents Specify components of the source bundle url.
 * @param devServerCommandLine If provided, starts the dev server.
 * @param devServerPort If provided, port to verify that the dev server is running.
 * @param packagerCommandLine If provided, starts the packager.
 * @param packagerHost Specifies the host name of the packager.
 * @param packagerPort Specifies the port of the packager.
 * @param sideloadCommandLine If provided, launches the add-in.
 * @param enableDebugging If false, start without debugging.
 */
export async function startDebugging(manifestPath: string,
    debuggingMethod: DebuggingMethod = defaultDebuggingMethod(),
    sourceBundleUrlComponents?: devSettings.SourceBundleUrlComponents,
    devServerCommandLine?: string, devServerPort?: number,
    packagerCommandLine?: string, packagerHost?: string, packagerPort?: string,
    sideloadCommandLine?: string, enableDebugging: boolean = true, enableLiveReload: boolean = true) {

    let packagerPromise: Promise<void> | undefined;
    let devServerPromise: Promise<void> | undefined;

    console.log(enableDebugging
        ? "Debugging is being started..."
        : "Starting without debugging...");

    if (packagerCommandLine) {
        packagerPromise = runPackager(packagerCommandLine, packagerHost, packagerPort);
    }

    if (devServerCommandLine) {
        devServerPromise = runDevServer(devServerCommandLine, devServerPort);
    }

    const manifestInfo = await manifest.readManifestFile(manifestPath);

    if (!manifestInfo.id) {
        throw new Error("Manifest does not contain the id for the Office Add-in.");
    }

    // enable debugging
    await devSettings.enableDebugging(manifestInfo.id, enableDebugging, debuggingMethod);
    if (enableDebugging) {
        console.log(`Enabled debugging for add-in ${manifestInfo.id}. Debug method: ${debuggingMethod.toString()}`);
    }

    // enable live reload
    await devSettings.enableLiveReload(manifestInfo.id, enableLiveReload);
    if (enableLiveReload) {
        console.log(`Enabled live-reload for add-in ${manifestInfo.id}.`);
    }

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

    if (enableDebugging && (debuggingMethod === DebuggingMethod.Web)) {
        try {
            await runNodeDebugger();
            console.log(`Started the node debugger.`);
        } catch (err) {
            console.log(`Unable to start the node debugger. ${err}`);
        }
    }

    if (sideloadCommandLine) {
        try {
            console.log(`Sideloading the Office Add-in...`);
            await startProcess(sideloadCommandLine);
        } catch (err) {
            console.log(`Unable to sideload the Office Add-in. ${err}`);
        }
    }

    console.log(enableDebugging
        ? "Debugging started."
        : "Started.");
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

export async function waitUntilDevServerIsRunning(port: number, retryCount: number = 10, retryDelay: number = 1000): Promise<boolean> {
    return waitUntil(() => {
        return isDevServerRunning(port);
    }, retryCount, retryDelay);
}

export async function waitUntilPackagerIsRunning(statusUrl: string, retryCount: number = 30, retryDelay: number = 1000): Promise<boolean> {
    return waitUntil(() => {
        return isPackagerRunning(statusUrl);
    }, retryCount, retryDelay);
}
