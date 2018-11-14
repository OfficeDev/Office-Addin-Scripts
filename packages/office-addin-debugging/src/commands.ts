import * as commander from "commander";
import * as devSettings from "office-addin-dev-settings";
import { parseDebuggingMethod, startDebugging} from "./start";
import { stopDebugging} from "./stop";

function logErrorMessage(err: any) {
    console.error(`Error: ${err instanceof Error ? err.message : err}`);
}

function parseDevServerPort(optionValue: any): number | undefined {
    const devServerPort = parseNumericCommandOption(optionValue, "--dev-server-port should specify a number.");

    if (devServerPort !== undefined) {
        if ((devServerPort < 0) || (devServerPort > 65535)) {
            throw new Error("--dev-server-port should be between 0 and 65535.");
        }
    }

    return devServerPort;
}

function parseNumericCommandOption(optionValue: any, errorMessage: string = "The value should be a number."): number | undefined {
    switch (typeof(optionValue)) {
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

export async function start(manifestPath: string, command: commander.Command) {
    try {
        const sourceBundleUrlComponents = new devSettings.SourceBundleUrlComponents(
            command.sourceBundleUrlHost, command.sourceBundleUrlPort,
            command.sourceBundleUrlPath, command.sourceBundleUrlExtension);
        const debuggingMethod = parseDebuggingMethod(command.debugMethod);
        const devServerPort = parseDevServerPort(command.devServerPort);

        await startDebugging(manifestPath, debuggingMethod, sourceBundleUrlComponents,
            command.devServer, devServerPort,
            command.packager, command.packagerHost, command.PackagerPort,
            command.sideload, command.debug, command.liveReload);
    } catch (err) {
        logErrorMessage(`Unable to start debugging.\n${err}`);
    }
}

export async function stop(manifestPath: string, command: commander.Command) {
    try {
        await stopDebugging(manifestPath, command.unload);
    } catch (err) {
        logErrorMessage(`Unable to stop debugging.\n${err}`);
    }
}
