// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as fs from "fs";
import * as defaults from "./defaults";
import { UsageDataLevel } from "./usageData";

/**
 * Allows developer to check if the program has already prompted before
 * @param groupName Group name of the usage data object
 * @returns Boolean of whether the program should prompt
 */
export function needToPromptForUsageData(groupName: string): boolean {
    return !groupNameExists(groupName);
}

/**
 * Reads data from the usage data json config file
 * @returns Parsed object from json file if it exists
 */
export function readUsageDataJsonData(): any {
    if (fs.existsSync(defaults.usageDataJsonFilePath)) {
        const jsonData = fs.readFileSync(defaults.usageDataJsonFilePath, "utf8");
        return JSON.parse(jsonData.toString());
    }
}
/**
 * Returns whether usage data is enabled on the usage data object
 * @param groupName Group name to search for in the specified json data
 * @returns Whether usage data is enabled specific to the group name
 */
export function readUsageDataLevel(groupName: string): UsageDataLevel {
    const jsonData = readUsageDataJsonData();
    return jsonData.usageDataInstances[groupName].usageDataLevel;
}
/**
 * Returns whether usage data is enabled on the usage data object
 * @param groupName Group name to search for in the specified json data
 * @param propertyName Property name that will be used to access and return the associated value
 * @returns Property of the specific group name
 */
export function readUsageDataObjectProperty(groupName: string, propertyName: string): any {
    const jsonData = readUsageDataJsonData();
    return jsonData.usageDataInstances[groupName][propertyName];
}
/**
 * Checks to see if the given group name exists in the specified json data
 * @param groupName Group name to search for in the specified json data
 * @returns Boolean of whether group name exists
 */
export function groupNameExists(groupName: string): boolean {
    if (fs.existsSync(defaults.usageDataJsonFilePath) && fs.readFileSync(defaults.usageDataJsonFilePath, "utf8") !== "" && fs.readFileSync(defaults.usageDataJsonFilePath, "utf8") !== "undefined") {
        const jsonData = readUsageDataJsonData();
        return Object.getOwnPropertyNames(jsonData.usageDataInstances).includes(groupName);
    }
    return false;
}
/**
 * Reads usage data settings from the usage data json config file for a specific group
 * @returns Settings for the specified group
 */
export function readUsageDataSettings(groupName = defaults.groupName): object | undefined {
    if (fs.existsSync(defaults.usageDataJsonFilePath)) {
        return readUsageDataJsonData().usageDataInstances[groupName];
    } else {
        return undefined;
    }
}