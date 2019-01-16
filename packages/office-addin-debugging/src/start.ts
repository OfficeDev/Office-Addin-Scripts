import * as fetch from "node-fetch";
import * as devSettings from "office-addin-dev-settings";
import { DebuggingMethod } from "office-addin-dev-settings";
import * as manifest from "office-addin-manifest";
import * as nodeDebugger from "office-addin-node-debugger";
import { getProcessIdsForPort } from "./port";
import { startDetachedProcess, startProcess  } from "./process";

export enum AppType {
    Desktop = "desktop",
    Web = "web",
}

function defaultDebuggingMethod(): DebuggingMethod {
    return DebuggingMethod.Direct;
}

function delay(milliseconds: number): Promise<void> {
    return new Promise<void>((resolve) => {
        setTimeout(resolve, milliseconds);
    });
}

export async function isDevServerRunning(port: number): Promise<boolean> {
    // isPortInUse(port) will return false when webpack-dev-server is running.
    // it should be fixed, but for now, use getProcessIdsForPort(port)
    const processIds = await getProcessIdsForPort(port);
    const isRunning = processIds.length > 0;

    return isRunning;
}

export async function isPackagerRunning(statusUrl: string): Promise<boolean> {
    const statusRunningResponse = `packager-status:running`;

    try {
        const response = await fetch.default(statusUrl);
        console.log(`packager: ${response.status} ${response.statusText}`);
        const text = await response.text();
        console.log(`packager: ${text}`);
        return (statusRunningResponse === text);
    } catch (err) {
        return false;
    }
}

export function parseAppType(text: string): AppType | undefined {
    switch (text) {
        case "desktop":
            return AppType.Desktop;
        case "web":
            return AppType.Web;
        default:
            return undefined;
    }
}

export function parseDebuggingMethod(text: string): DebuggingMethod | undefined {
    switch (text) {
        case "direct":
            return DebuggingMethod.Direct;
        case "proxy":
            return DebuggingMethod.Proxy;
        default:
            return undefined;
    }
}

export async function runDevServer(commandLine: string, port?: number): Promise<void> {
    if (commandLine) {
        // if the dev server is running
        if ((port !== undefined) && await isDevServerRunning(port)) {
            console.log(`The dev server is already running on port ${port}.`);
        } else {
            // start the dev server
            console.log(`Starting the dev server... (${commandLine})`);
            startDetachedProcess(commandLine);

            if (port !== undefined) {
                // wait until the dev server is running
                const isRunning: boolean = await waitUntilDevServerIsRunning(port);

                if (isRunning) {
                    console.log(`The dev server is running on port ${port}.`);
                } else {
                    throw new Error(`The dev server is not running on port ${port}.`);
                }
            }
        }
    }
}

export async function runNodeDebugger(host?: string, port?: string): Promise<void> {
    nodeDebugger.run(host, port);

    console.log("The node debugger is running.");
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
            await startDetachedProcess(commandLine);

            // wait until the packager is running
            if (await waitUntilPackagerIsRunning(statusUrl)) {
                console.log(`The packager is running. ${packagerUrl}`);
            } else {
                throw new Error(`The packager is not running. ${packagerUrl}`);
            }
        }
    }
}

/**
 * Start debugging
 * @param appType The type of application to debug.
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
export async function startDebugging(manifestPath: string, appType: AppType,
    debuggingMethod: DebuggingMethod = defaultDebuggingMethod(),
    sourceBundleUrlComponents?: devSettings.SourceBundleUrlComponents,
    devServerCommandLine?: string, devServerPort?: number,
    packagerCommandLine?: string, packagerHost?: string, packagerPort?: string,
    sideloadCommandLine?: string, enableDebugging: boolean = true, enableLiveReload: boolean = (debuggingMethod === DebuggingMethod.Proxy)) {

    const isDesktopAppType = (appType === AppType.Desktop);
    const isWebAppType = (appType === AppType.Web);
    const isProxyDebuggingMethod = (debuggingMethod === DebuggingMethod.Proxy);
    let packagerPromise: Promise<void> | undefined;
    let devServerPromise: Promise<void> | undefined;

    console.log(enableDebugging
        ? "Debugging is being started..."
        : "Starting without debugging...");
    console.log(`App type: ${appType.toString()}`);

    const manifestInfo = await manifest.readManifestFile(manifestPath);

    if (!manifestInfo.id) {
        throw new Error("Manifest does not contain the id for the Office Add-in.");
    }

    // enable debugging
    if (isDesktopAppType) {
        await devSettings.enableDebugging(manifestInfo.id, enableDebugging, debuggingMethod);
        if (enableDebugging) {
            console.log(`Enabled debugging for add-in ${manifestInfo.id}. Debug method: ${debuggingMethod.toString()}`);
        }
    }

    // enable live reload
    if (isDesktopAppType) {
        await devSettings.enableLiveReload(manifestInfo.id, enableLiveReload);
        if (enableLiveReload) {
            console.log(`Enabled live-reload for add-in ${manifestInfo.id}.`);
        }
    }

    // set source bundle url
    if (isDesktopAppType) {
        if (sourceBundleUrlComponents) {
            await devSettings.setSourceBundleUrl(manifestInfo.id, sourceBundleUrlComponents);
        }
    }

    if (packagerCommandLine && isProxyDebuggingMethod && isDesktopAppType) {
        packagerPromise = runPackager(packagerCommandLine, packagerHost, packagerPort);
    }

    if (devServerCommandLine) {
        devServerPromise = runDevServer(devServerCommandLine, devServerPort);
    }

    if (packagerPromise !== undefined) {
        try {
            await packagerPromise;
        } catch (err) {
            console.log(`Unable to start the packager. ${err}`);
        }
    }

    if (devServerPromise !== undefined) {
        try {
            await devServerPromise;
        } catch (err) {
            console.log(`Unable to start the dev server. ${err}`);
        }
    }

    if (enableDebugging && isProxyDebuggingMethod && isDesktopAppType) {
        try {
            await runNodeDebugger();
        } catch (err) {
            console.log(`Unable to start the node debugger. ${err}`);
        }
    }

    if (sideloadCommandLine && isDesktopAppType) {
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
        await delay(retryDelay);
        done = await callback();
    }

    return done;
}

export async function waitUntilDevServerIsRunning(port: number, retryCount: number = 30, retryDelay: number = 1000): Promise<boolean> {
    return waitUntil(async () => await isDevServerRunning(port), retryCount, retryDelay);
}

export async function waitUntilPackagerIsRunning(statusUrl: string, retryCount: number = 30, retryDelay: number = 1000): Promise<boolean> {
    return waitUntil(async () => await isPackagerRunning(statusUrl), retryCount, retryDelay);
}
