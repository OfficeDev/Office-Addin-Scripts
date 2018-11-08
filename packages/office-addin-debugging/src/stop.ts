import * as devSettings from "office-addin-dev-settings";
import * as manifest from "office-addin-manifest";
import { startProcess } from "./process";

export async function stopDebugging(manifestPath: string, unregisterCommandLine?: string) {
    console.log("Debugging is being stopped...");

    const manifestInfo = await manifest.readManifestFile(manifestPath);

    if (!manifestInfo.id) {
        throw new Error("Manifest does not contain the id for the Office Add-in.");
    }

    // clear dev settings
    await devSettings.clearDevSettings(manifestInfo.id);

    if (unregisterCommandLine) {
        // unregister
        try {
            await startProcess(unregisterCommandLine);
        } catch (err) {
            console.log(`Unable to unregister the Office Add-in. ${err}`);
        }
    }

    console.log("Debugging has being stopped.");
}
