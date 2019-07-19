import * as fs from "fs";
import * as os from "os";
import { telemetryObject } from "./officeAddinTelemetry";
const optInResponsePath: string = os.homedir() + "/officeAddinTelemetry.json";

/**
 * Write the telemetryObject info to the json file
 * @param telemetryObject DebuggingInfo object
 */
export function writeTelemetryInfo(telemetryObject: telemetryObject, pathToFile: string) {
    const debuggingInfoJSON = JSON.stringify(telemetryObject, null, 4);
    fs.writeFileSync(optInResponsePath, debuggingInfoJSON);
}

/**
 * Read the DebuggingInfo object
 */
export function readTelemetyInfo(): any {
    if (fs.existsSync(optInResponsePath)) {
    const json = fs.readFileSync(optInResponsePath);
    return JSON.parse(json.toString());
    }
}
