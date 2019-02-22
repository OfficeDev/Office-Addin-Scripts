import * as commnder from "commander";
import { startTestServer, stopTestServer }  from "./testServer"

export async function start(command: commnder.Command) {
    const port: string | undefined = getCommandOptionString(command.port, "");
    const serverStarted: boolean  = await startTestServer(port);

    if (serverStarted){
        console.log("Server started successfully");
    }
    else{
        console.log("Server failed to start");
    }
}

export async function stop() {
    const serverStopped: boolean = await stopTestServer();
    if (serverStopped) {
        console.log("Server stopped successfully");
    }
    else {
        console.log("Server failed to stop");
    }
}

function getCommandOptionString(option: string | boolean, defaultValue?: string): string | undefined {
    // For a command option defined with an optional value, e.g. "--option [value]",
    // when the option is provided with a value, it will be of type "string", return the specified value;
    // when the option is provided without a value, it will be of type "boolean", return undefined.
    return (typeof (option) === "boolean") ? defaultValue : option;
}
