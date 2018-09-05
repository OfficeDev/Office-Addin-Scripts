#!/usr/bin/env node

// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as fetch from 'node-fetch';
import * as utils from './util';
import * as commander from 'commander';
import { exec } from 'child_process';
import { execSync } from 'child_process';
import { ExecException } from 'child_process';

function handleCommandError(error: ExecException | null, stdout: string, stderr: string): void {
    if (error) {
        console.log(`${error}`);
        return;
    }
    console.log(`stdout: ${stdout}`);
    console.log(`stderr: ${stderr}`);
}

commander
    .option('-h, --hostname <hostname>', 'The hostname where the packager is running.')
    .option('-p, --port <port>', 'The port where the packager is running.')
    .option('--dev-server', 'Flag that indicates whether you want to run the dev server or not.')
    .parse(process.argv);

const hostName: string = commander.hostname ? commander.hostname : 'localhost';
const port: string = commander.port ? commander.port : '8081';
const devServer: boolean = commander.devServer ? commander.devServer : false;
const bundlerUrl: string = `http://${hostName}:${port}/status`;
const bundleRunningText = `packager-status:running`;
const fetchRetry: number = 10;
const fetchRetryTimeout: number = 1000;

//This is a temporal "shortcut" for now, we should create an independent package which handles the packager run. Also, we assume that the packager scripts are in place.
let haulCmd = `npm run packager`;
exec(haulCmd, handleCommandError);

//Wait until haul server is started successfully.
utils.retryFetch(bundlerUrl, {}, fetchRetry, fetchRetryTimeout).then((res: fetch.Response) => {
    if (res.ok) {
        res.text().then(
            text => {
                if (bundleRunningText.match(text)) {
                    console.log(`The packager is up and running! ${bundlerUrl}`);

                    if (devServer) {
                        //Start the dev server, this is leveraging from the "dev-server" script in the package.json, that script name could change so we should have a better way for invoking it in the future.
                        console.log(`Starting dev-server...`);
                        let devServerCmd = `npm run dev-server`;
                        exec(devServerCmd, handleCommandError)
                    }

                    console.log(`Initializing Debugger...`);
                    let debuggerCmd = `office-addin-node-debugger -p ${port} -h ${hostName}`;
                    exec(debuggerCmd, handleCommandError);

                    console.log(`Sideloading the app...`);
                    let sideloadCmd = `npm run sideload`;
                    execSync(sideloadCmd);
                }
            },
            error => {
                console.log(`The packager failed to run. ${error}`);
            }
        );
    } else {
        console.log(`Unable to connect to packager. Status: ${res.status}  StatusText:${res.statusText}`);
    }
});