
import * as chalk from "chalk";
import * as fs from "fs";
import * as os from "os";
import * as path from "path";
import { telemetryLevel } from "./officeAddinTelemetry";
const telemetryJsonFilePath: string = path.join(os.homedir(), "/officeAddinTelemetry.json");

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
 * Allows developer to add or modify a specific property to the group
 * @param groupName Group name of property
 * @param property Property that will be created or modified
 * @param value Property's value that will be assigned
 * @param jsonFilePath Optional path to the json config file
 */
export function modifyTelemetryJsonData(groupName: string, property: any, value: any, jsonFilePath = telemetryJsonFilePath): void {
    try {
        if (fs.existsSync(jsonFilePath)) {
            const telemetryJsonData = readTelemetryJsonData(jsonFilePath);
            if (groupNameExists(telemetryJsonData, groupName)) {
                telemetryJsonData.telemetryInstances[groupName][property] = value;
                fs.writeFileSync(jsonFilePath, JSON.stringify((telemetryJsonData), null, 2));
            }
        }
    } catch {
        console.log(chalk.default.red("Could not modify property!\n Make sure the group name and file exists!"));
    }
}
/**
 * Reads data from the telemetry json config file
 * @param jsonFilePath Optional path to the json config file
 * @returns Parsed object from json file if it exists
 */
export function readTelemetryJsonData(jsonFilePath = telemetryJsonFilePath): any {
    if (fs.existsSync(jsonFilePath)) {
        const jsonData = fs.readFileSync(jsonFilePath, "utf8");
        return JSON.parse(jsonData.toString());
    }
}
/**
 * Returns telemetry level of telemetry object
 * @param groupName Group name to search for in the specified json data
 * @param jsonFilePath Optional path to the json config file
 */
export function readTelemetryLevel(groupName: string, jsonFilePath = telemetryJsonFilePath): string {
    const jsonData = readTelemetryJsonData(jsonFilePath);
    return jsonData.telemetryInstances[groupName].telemetryLevel;
}
/**
 * Returns telemetry property of telemetry object
 * @param groupName Group name to search for in the specified json data
 * @param propertyName Property name that will be used to access and return the associated value
 * @param jsonFilePath Optional path to the json config file
 */
export function readTelemetryObjectProperty(groupName: string, propertyName: string, jsonFilePath = telemetryJsonFilePath): string {
    const jsonData = readTelemetryJsonData(jsonFilePath);
    return jsonData.telemetryInstances[groupName][propertyName];
}
/**
 * Writes to telemetry config file either appending to already existing file or creating new file
 * @param groupName Group name of telemetry object
 * @param telemetryLevel Whether user is sending basic or verbose telemetry
 * @param jsonFilePath Optional path to the json config file
 */

export function writeTelemetryJsonData(groupName: string, level: telemetryLevel, jsonFilePath = telemetryJsonFilePath): void {
    if (fs.existsSync(jsonFilePath) && fs.readFileSync(jsonFilePath, "utf8") !== "") {
        const telemetryJsonData = readTelemetryJsonData(jsonFilePath);
        if (groupNameExists(telemetryJsonData, groupName)) {
            modifyTelemetryJsonData(groupName, "telemetryLevel", level);
        } else {
            telemetryJsonData.telemetryInstances[groupName] = { telemetryLevel: String };
            telemetryJsonData.telemetryInstances[groupName].telemetryLevel = level;
            fs.writeFileSync(jsonFilePath, JSON.stringify((telemetryJsonData), null, 2));
        }
    } else {
        let jsonData = {};
        jsonData[groupName] = telemetryLevel;
        jsonData = { telemetryInstances: jsonData };
        jsonData = { telemetryInstances: { [groupName]: { telemetryLevel } } };
        fs.writeFileSync(jsonFilePath, JSON.stringify((jsonData), null, 2));
    }
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
