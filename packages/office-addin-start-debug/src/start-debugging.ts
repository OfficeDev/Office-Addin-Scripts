#!/usr/bin/env node

// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as fetch from 'node-fetch';
import * as commander from 'commander';
import { exec } from 'child_process';
import { execSync } from 'child_process';
import { ExecException } from 'child_process';
import { DebuggingMethod } from 'office-addin-dev-settings';

function delay(milliseconds: number): Promise<void> {
    return new Promise<void>(resolve => {
        setTimeout(resolve, milliseconds);
    });
}

function defaultDebuggingMethod(): DebuggingMethod {
    switch (process.platform) {
        case "win32":
            return DebuggingMethod.Web; // will change to Direct later 
        default:
            return DebuggingMethod.Web;
    }
}

export async function isDevServerRunning(url: string): Promise<boolean> {
    try {
        console.log(`Is dev server running? (${url})`)
        const response = await fetch.default(url);
        console.log(`devServer: ${response.status} ${response.statusText}`)
        return response.ok;
    } catch (err) {
        console.log(`Unable to connect to dev server.\n${err}`);
        return false;
    }    
}

export async function isPackagerRunning(statusUrl: string): Promise<boolean> {
    const statusRunningResponse = `packager-status:running`;

    try {
        console.log(`Is packager running? (${statusUrl})`)
        const response = await fetch.default(statusUrl);
        console.log(`packager: ${response.status} ${response.statusText}`)
        const text = await response.text();
        console.log(`packager: ${text}`)
        return (statusRunningResponse === text);
    } catch (err) {
        return false;
    }    
}

async function startProcess(commandLine: string, async: boolean = true): Promise<void> {
    if (async) {
        return new Promise<void>((resolve, reject) => {
            console.log(`Starting process: ${commandLine}`);
            exec(commandLine, (error: ExecException | null, stdout: string, stderr: string) => {
                if (error) {
                    reject(error);
                } else {
                    resolve();
                }
            });
        });
    } else {
        execSync(commandLine);
        Promise.resolve();
    }
}

export async function runDevServer(commandLine: string, url?: string): Promise<void> {
    if (commandLine) {
        // if the dev server is running
        if (url && await isDevServerRunning(url)) {
            console.log(`The dev server is already running. ${url}`);
        } else {
            // start the packager
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
    console.log(`Initializing Debugger...`);
    let debuggerCmd = `office-addin-node-debugger`;
    
    if (host) {
        debuggerCmd += `-h ${host}`;
    }

    if (port) {
        debuggerCmd += `-p ${port}`;
    }

    return startProcess(debuggerCmd);
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
        return isDevServerRunning(url) 
    }, retryCount, retryDelay);
}

export async function waitUntilPackagerIsRunning(statusUrl: string, retryCount: number = 30, retryDelay: number = 1000): Promise<boolean> {
    return waitUntil(() => { 
        return isPackagerRunning(statusUrl)
    }, retryCount, retryDelay);
}

export async function startDebugging(method: DebuggingMethod = defaultDebuggingMethod(), 
    devServerCommandLine?: string, devServerUrl?: string, 
    packagerCommandLine?: string, packagerHost?: string, packagerPort?: string,
    sideloadCommandLine: string = "npm run sideload") {
    let packagerPromise: Promise<void> = Promise.resolve();
    let devServerPromise: Promise<void> = Promise.resolve();
    let nodeDebuggerPromise: Promise<void> = Promise.resolve();

    if (packagerCommandLine) {
        packagerPromise = runPackager(packagerCommandLine, packagerHost, packagerPort);
    }

    if (devServerCommandLine) {
        let devServerPromise = runDevServer(devServerCommandLine, devServerUrl || "http://localhost:8081");
    } 

    nodeDebuggerPromise = runNodeDebugger();

    try {
        await packagerPromise;
    } catch (err) {
        console.log(`Unable to start the packager. ${err}`);
    }

    try {
        await devServerPromise;
    } catch (err) {
        console.log(`Unable to start the dev server. ${err}`);
    }

    try {
        await nodeDebuggerPromise;
    } catch (err) {
        console.log(`Unable to start the node debugger. ${err}`);
    }
    
    console.log(`Sideloading the app...`);
    await startProcess(sideloadCommandLine);
}


if (process.argv[1].endsWith("\\start-debugging.js")) {
    commander
        .option('-h, --host <host>', 'The host where the packager is running.')
        .option('-p, --port <port>', 'The port where the packager is running.')
        .option('--debug-method <method>', 'The debug method to use')
        .option('--dev-server [command]', 'Run the dev server.')
        .option('--dev-server-url <url>')
        .option('--packager [command]', 'Run the packager')
        .parse(process.argv);

    try {
        startDebugging(commander.debugMethod, commander.devServer, commander.devServerUrl, commander.packager);
    } catch (err) {
        console.log(`Unable to start debugging.\n${err}`);
    }
}
