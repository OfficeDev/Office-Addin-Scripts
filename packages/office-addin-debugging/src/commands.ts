import * as commander from "commander";
import * as devSettings from "office-addin-dev-settings";
import { parseDebuggingMethod, startDebugging} from "./start";
import { stopDebugging} from "./stop";

export async function start(manifestPath: string, command: commander.Command) {
    try {
        const sourceBundleUrlComponents = new devSettings.SourceBundleUrlComponents(
            command.sourceBundleUrlHost, command.sourceBundleUrlPort,
            command.sourceBundleUrlPath, command.sourceBundleUrlExtension);
        const debuggingMethod = parseDebuggingMethod(command.debugMethod);

        startDebugging(manifestPath, debuggingMethod, sourceBundleUrlComponents,
            command.devServer, command.devServerPort,
            command.packager, command.packagerHost, command.PackagerPort,
            command.sideload, command.debug, command.liveReload);
    } catch (err) {
        console.log(`Unable to start debugging.\n${err}`);
    }
}

export async function stop(manifestPath: string, command: commander.Command) {
    try {
        stopDebugging(manifestPath, command.unload);
    } catch (err) {
        console.log(`Unable to stop debugging.\n${err}`);
    }
}
