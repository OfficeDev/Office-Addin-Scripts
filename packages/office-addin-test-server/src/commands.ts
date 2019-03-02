import * as commnder from "commander";
import TestServer from "./testServer"
const port: number = 8080


export async function start(command: commnder.Command) {
    const port: string | undefined = getCommandOptionString(command.port, "");
    let testServerPort: number | undefined;

    if (port != undefined){
        testServerPort = parseTestServerPort(port);
    }

    const testServer = new TestServer(testServerPort);
    const serverStarted: boolean = await testServer.startTestServer();

    if (serverStarted){
        console.log(`Server started successfully on port ${testServerPort != undefined ? testServerPort: 8080}`);
    }
    else{
        console.log("Server failed to start");
    }
}

function getCommandOptionString(option: string | boolean, defaultValue?: string): string | undefined {
    // For a command option defined with an optional value, e.g. "--option [value]",
    // when the option is provided with a value, it will be of type "string", return the specified value;
    // when the option is provided without a value, it will be of type "boolean", return undefined.
    return (typeof (option) === "boolean") ? defaultValue : option;
}

function parseNumericCommandOption(optionValue: any, errorMessage: string = "The value should be a number."): number | undefined {
    switch (typeof (optionValue)) {
        case "number": {
            return optionValue;
        }
        case "string": {
            let result;

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

function parseTestServerPort(optionValue: any): number | undefined {
    const testServerPort = parseNumericCommandOption(optionValue, "--dev-server-port should specify a number.");

    if (testServerPort !== undefined) {
        if ((testServerPort < 0) || (testServerPort > 65535)) {
            throw new Error("port should be between 0 and 65535.");
        }
    }
    return testServerPort;
}