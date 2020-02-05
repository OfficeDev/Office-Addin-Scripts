// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as usageData from "office-addin-usage-data";

export const lintFiles = "src/**/*.{ts,tsx,js,jsx}";

// Usage data defaults
export const usageDataProjectName: string = "office-addin-lint";
export const sendUsageData: boolean = usageData.groupNameExists(usageData.groupName) && usageData.readUsageDataLevel(usageData.groupName) === usageData.UsageDataLevel.on;
export const usageDataOptions: usageData.IUsageDataOptions = {
    groupName: usageData.groupName,
    projectName: usageDataProjectName,
    raisePrompt: false,
    instrumentationKey: usageData.instrumentationKeyForOfficeAddinCLITools,
    promptQuestion: "",
    usageDataLevel: sendUsageData ? usageData.UsageDataLevel.on : usageData.UsageDataLevel.off,
    method: usageData.UsageDataReportingMethod.applicationInsights,
    isForTesting: true
};
const usageDataObject: usageData.OfficeAddinUsageData = new usageData.OfficeAddinUsageData(usageDataOptions);
export function sendUsageDataSuccessEvent(method: string, ...data: object[]) {
    usageDataObject.sendUsageDataSuccessEvent(usageDataProjectName, data.concat({Method: method}));
};
export function sendUsageDataException(method: string, error: Error | string, ...data: Object[]) {
    usageDataObject.sendUsageDataException(usageDataProjectName, error, data.concat({Method: method}));
};
export function sendUsageDataCustomEvent(usageDataInfo: Object) {
    usageDataObject.sendUsageDataEvent(usageDataProjectName, usageDataInfo);
};