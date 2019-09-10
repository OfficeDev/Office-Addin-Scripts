// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as commnder from "commander";
import { parseNumber } from "office-addin-cli";
import * as defaults from "./defaults";
import { TestServer } from "./testServer";

export async function start(command: commnder.Command) {
    const httpPort: number = (command.http !== undefined) ? parseTestServerPort(command.http) : defaults.httpPort;
    const httpsPort: number = (command.https !== undefined) ? parseTestServerPort(command.https) : defaults.httpsPort;
    const testServer = new TestServer(httpsPort, httpPort);
    const serverStarted: boolean = await testServer.startTestServer();

    if (serverStarted) {
        console.log(`Server started successfully on port ${httpsPort} and ${httpPort}`);
    } else {
        console.log("Server failed to start");
    }
}

function parseNumericCommandOption(optionValue: any, errorMessage: string = "The value should be a number."): number {
    switch (typeof (optionValue)) {
        case "number": {
            return optionValue;
        }
        case "string": {
            let result: number;

            try {
                result = parseInt(optionValue, 10);
            } catch (err) {
                throw new Error(errorMessage);
            }

            if (Number.isNaN(result)) {
                throw new Error(errorMessage);
            }

            return result;
        }
        case "undefined": {
            return undefined;
        }
        default: {
            throw new Error(errorMessage);
        }
    }
}

function parseTestServerPort(optionValue: any): number {
    const testServerPort = parseNumber(optionValue, "--dev-server-port should specify a number.");

    if (testServerPort !== undefined) {
        if ((testServerPort < 0) || (testServerPort > 65535)) {
            throw new Error("port should be between 0 and 65535.");
        }
    }
    return testServerPort;
}
