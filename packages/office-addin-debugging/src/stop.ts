import * as devSettings from "office-addin-dev-settings";
import * as manifest from "office-addin-manifest";
import { startProcess, stopProcess } from "./process";

export async function stopDebugging(manifestPath: string, unregisterCommandLine?: string) {
    console.log("Debugging is being stopped...");

    const isWindowsPlatform = (process.platform === "win32");
    const manifestInfo = await manifest.readManifestFile(manifestPath);

    if (!manifestInfo.id) {
        throw new Error("Manifest does not contain the id for the Office Add-in.");
    }

    // clear dev settings
    if (isWindowsPlatform) {
      await devSettings.clearDevSettings(manifestInfo.id);
    }

    if (unregisterCommandLine) {
        // unregister
        try {
            await startProcess(unregisterCommandLine);
        } catch (err) {
            console.log(`Unable to unregister the Office Add-in. ${err}`);
        }
    }

    if (process.env.OfficeAddinDevServerProcessId) {
        stopProcess();
        delete process.env.OfficeAddinDevServerProcessId;
    }

    console.log("Debugging has been stopped.");
}
