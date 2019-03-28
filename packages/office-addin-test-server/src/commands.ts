import * as commnder from "commander";
import { TestServer } from "./testServer"
const defaultPort: number = 4201;

export async function start(command: commnder.Command) {
    const testServerPort: number = (command.port !== undefined) ? parseTestServerPort(command.port) : defaultPort;
    const testServer = new TestServer(testServerPort);
    const serverStarted: boolean = await testServer.startTestServer();

    if (serverStarted){
        console.log(`Server started successfully on port ${testServerPort}`);
    }
    else{
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
    const testServerPort = parseNumericCommandOption(optionValue, "--dev-server-port should specify a number.");

    if (testServerPort !== undefined) {
        if ((testServerPort < 0) || (testServerPort > 65535)) {
            throw new Error("port should be between 0 and 65535.");
        }
    }
    return testServerPort;
}