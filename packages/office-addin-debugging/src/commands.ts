import * as commander from "commander";
import * as devSettings from "office-addin-dev-settings";
import * as debugging from "./debugging";

export async function start(manifestPath: string, command: commander.Command) {

    try {
        const sourceBundleUrlComponents = new devSettings.SourceBundleUrlComponents(
            command.sourceBundleUrlHost, command.sourceBundleUrlPort,
            command.sourceBundleUrlPath, command.sourceBundleUrlExtension);
        const debuggingMethod = debugging.parseDebuggingMethod(command.debugMethod);

        debugging.startDebugging(manifestPath, debuggingMethod, sourceBundleUrlComponents,
            command.devServer, command.devServerUrl, command.packager);
    } catch (err) {
        console.log(`Unable to start debugging.\n${err}`);
    }
}
