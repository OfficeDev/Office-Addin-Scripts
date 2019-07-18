import * as fs from "fs";
import {  telemetryObject} from "./officeAddinTelemetry";
/**
 * Write the DebuggingInfo to the json file
 * @param debuggingInfo DebuggingInfo object
 * @param pathToFile Path to the json file containing debug info
 */
export function writeDebuggingInfo(debuggingInfo: telemetryObject, pathToFile: string) {
    const debuggingInfoJSON = JSON.stringify(debuggingInfo, null, 4 );
    fs.writeFileSync(pathToFile, debuggingInfoJSON);
}

/**
 * Read the DebuggingInfo object
 * @param pathToFile Path to the json file containing debug info
 */
export function readDebuggingInfo(pathToFile: string): any {
    const json = fs.readFileSync(pathToFile);
    return JSON.parse(json.toString());
}

/**
 * Reports custom event object to telemetry structure
 * @param eventName Event name sent to telemetry structure
 * @param data Data object sent to telemetry structure
 * @param timeElapsed Optional parameter for custom metric in data object sent
 */
export function reportEvent(eventName: string, data: object, timeElapsed = 0) {
}

/**
 * Reports custom event object to telemetry structure
 * @param errorName Error name sent to telemetry structure
 * @param err Error sent to telemetry structure
 */
export function reportError(errorName: string, err: Error) {

}