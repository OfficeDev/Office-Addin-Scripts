
import * as chalk from "chalk";
import * as fs from "fs";
import * as os from "os";
import * as path from "path";
/**
 * Allows developer to create prompts and responses in other applications before object creation
 * @param groupName Event name sent to telemetry structure
 * @param telemetryEnabled Whether user agreed to data collection
 * @returns Boolean of whether the program should prompt
 */
export function promptForTelemetry(groupName: string, jsonFilePath): boolean {
    try {
        const jsonData: any = readTelemetryJsonData(jsonFilePath);
        if (jsonData) {
            return !groupNameExists(jsonData, groupName);
        }
        return true;
    } catch (err) {
        console.log(chalk.default.red(err));
    }
}

/**
 * Reads data from the telemetry json config file
 * @param jsonFilePath Path to the json config file
 * @returns Parsed object from json file if it exists
 */
export function readTelemetryJsonData(jsonFilePath: string): any {
    if (fs.existsSync(jsonFilePath)) {
        const jsonData = fs.readFileSync(jsonFilePath, "utf8");
        return JSON.parse(jsonData.toString());
    }
}

/**
 * Writes data to the telemetry json config file
 * @param jsonData Telemetry json data to write to the json config file
 * @param jsonFilePath Path to the json config file
 */
export function writeTelemetryJsonData(jsonData: any, jsonFilePath = path.join(os.homedir(), "/officeAddinTelemetry.json")): void {
    fs.writeFileSync(jsonFilePath, JSON.stringify((jsonData), null, 2));
}

/**
 * Writes new telemetry json config file if one doesn't already exist
 * @param groupName Telemetry group name to write to the json config file
 * @param telemetryEnabled Specifies whether opted into telemetry collection
 * @param jsonFilePath Path to the json config file
 */
export function writeNewTelemetryJsonFile(groupName: string, telemetryEnabled: boolean, jsonFilePath = path.join(os.homedir(), "/officeAddinTelemetry.json")): void {
    let jsonData = {};
    jsonData[groupName] = telemetryEnabled;
    jsonData = { telemetryInstances: jsonData};
    jsonData = { telemetryInstances: {[groupName]: {telemetryEnabled}} };
    writeTelemetryJsonData(jsonData, jsonFilePath);
}

/**
 * Checks to see if a give group name exists in the specified json data
 * @param jsonData Telemetry json data to search
 * @param groupName Group name to search for in the specified json data
 * @returns Boolean of whether group name exists
 */
export function groupNameExists(jsonData: any, groupName: string): boolean {
    return Object.getOwnPropertyNames(jsonData.telemetryInstances).includes(groupName);
}
