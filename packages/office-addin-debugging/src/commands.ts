import * as commander from "commander";
import * as devSettings from "office-addin-dev-settings";
import { parseDebuggingMethod, startDebugging} from "./start";

export async function start(manifestPath: string, command: commander.Command) {
    try {
        const sourceBundleUrlComponents = new devSettings.SourceBundleUrlComponents(
            command.sourceBundleUrlHost, command.sourceBundleUrlPort,
            command.sourceBundleUrlPath, command.sourceBundleUrlExtension);
        const debuggingMethod = parseDebuggingMethod(command.debugMethod);

        startDebugging(manifestPath, debuggingMethod, sourceBundleUrlComponents,
            command.devServer, command.devServerUrl,
            command.packager, command.packagerHost, command.PackagerPort);
    } catch (err) {
        console.log(`Unable to start debugging.\n${err}`);
    }
}
