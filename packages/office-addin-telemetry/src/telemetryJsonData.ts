
import * as chalk from "chalk";
import * as fs from "fs";
import * as os from "os";
import * as path from "path";
import { telemetryLevel } from "./officeAddinTelemetry";
const telemetryJsonFilePath: string = path.join(os.homedir(), "/officeAddinTelemetry.json");

/**
 * Allows developer to check to check if the program has already prompted before
 * @param groupName Group name of the telemetry object
 * @param jsonFilePath Optional path to the json config file
 * @returns Boolean of whether the program should prompt
 */
export function promptForTelemetry(groupName: string, jsonFilePath: string = telemetryJsonFilePath): boolean {
    try {
        return !groupNameExists(groupName, jsonFilePath);
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
export function modifyTelemetryJsonData(groupName: string, property: any, value: any, jsonFilePath: string = telemetryJsonFilePath): void {
    try {
        if (fs.existsSync(jsonFilePath)) {
            if (groupNameExists(groupName, jsonFilePath)) {
                const telemetryJsonData = readTelemetryJsonData(jsonFilePath);
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
export function readTelemetryJsonData(jsonFilePath: string = telemetryJsonFilePath): any {
    if (fs.existsSync(jsonFilePath)) {
        const jsonData = fs.readFileSync(jsonFilePath, "utf8");
        return JSON.parse(jsonData.toString());
    }
}
/**
 * Returns telemetry level of telemetry object
 * @param groupName Group name to search for in the specified json data
 * @param jsonFilePath Optional path to the json config file
 * @returns Telemetry level specific to the group name
 */
export function readTelemetryLevel(groupName: string, jsonFilePath: string = telemetryJsonFilePath): telemetryLevel {
    const jsonData = readTelemetryJsonData(jsonFilePath);
    return jsonData.telemetryInstances[groupName].telemetryLevel;
}
/**
 * Returns telemetry property of telemetry object
 * @param groupName Group name to search for in the specified json data
 * @param propertyName Property name that will be used to access and return the associated value
 * @param jsonFilePath Optional path to the json config file
 * @returns Property of the specific group name
 */
export function readTelemetryObjectProperty(groupName: string, propertyName: string, jsonFilePath: string = telemetryJsonFilePath): any {
    const jsonData = readTelemetryJsonData(jsonFilePath);
    return jsonData.telemetryInstances[groupName][propertyName];
}
/**
 * Writes to telemetry config file either appending to already existing file or creating new file
 * @param groupName Group name of telemetry object
 * @param telemetryLevel Whether user is sending basic or verbose telemetry
 * @param jsonFilePath Optional path to the json config file
 */

export function writeTelemetryJsonData(groupName: string, level: telemetryLevel, jsonFilePath: string = telemetryJsonFilePath): void {
    if (fs.existsSync(jsonFilePath) && fs.readFileSync(jsonFilePath, "utf8") !== "") {
        if (groupNameExists(groupName, jsonFilePath)) {
            modifyTelemetryJsonData(groupName, "telemetryLevel", level);
        } else {
            const telemetryJsonData = readTelemetryJsonData(jsonFilePath);
            telemetryJsonData.telemetryInstances[groupName] = { telemetryLevel: String };
            telemetryJsonData.telemetryInstances[groupName].telemetryLevel = level;
            fs.writeFileSync(jsonFilePath, JSON.stringify((telemetryJsonData), null, 2));
        }
    } else {
        let telemetryJsonData = {};
        telemetryJsonData[groupName] = level;
        telemetryJsonData = { telemetryInstances: telemetryJsonData };
        telemetryJsonData = { telemetryInstances: { [groupName]: { ["telemetryLevel"]: level } } };
        fs.writeFileSync(jsonFilePath, JSON.stringify((telemetryJsonData), null, 2));
    }
}
/**
 * Checks to see if the given group name exists in the specified json data
 * @param groupName Group name to search for in the specified json data
 * @returns Boolean of whether group name exists
 */
export function groupNameExists(groupName: string, jsonFilePath: string = telemetryJsonFilePath): boolean {
    if (fs.existsSync(jsonFilePath) && fs.readFileSync(jsonFilePath, "utf8") !== "") {
        const jsonData = readTelemetryJsonData(jsonFilePath);
        return Object.getOwnPropertyNames(jsonData.telemetryInstances).includes(groupName);
    }
    return false;
}
