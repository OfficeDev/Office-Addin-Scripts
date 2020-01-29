// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as os from "os";
import * as path from "path";
import * as usageData from "office-addin-usage-data";

// Default certificate names
export const certificateDirectoryName = ".office-addin-dev-certs";
export const certificateDirectory =  path.join(os.homedir(), certificateDirectoryName);
export const caCertificateFileName = "ca.crt";
export const caCertificatePath = path.join(certificateDirectory, caCertificateFileName);
export const localhostCertificateFileName = "localhost.crt";
export const localhostCertificatePath = path.join(certificateDirectory, localhostCertificateFileName);
export const localhostKeyFileName = "localhost.key";
export const localhostKeyPath = path.join(certificateDirectory, localhostKeyFileName);

// Default certificate details
export const certificateName = "Developer CA for Microsoft Office Add-ins";
export const countryCode = "US";
export const daysUntilCertificateExpires = 30;
export const domain = ["127.0.0.1", "localhost"];
export const locality = "Redmond";
export const state = "WA";

// Usage data defaults
export const usageDataProjectName: string = "office-addin-dev-certs";
export const sendUsageData: boolean = usageData.groupNameExists(usageData.groupName) && usageData.readUsageDataLevel(usageData.groupName) === usageData.UsageDataLevel.on;
export const usageDataOptions: usageData.IUsageDataOptions = {
    groupName: usageData.groupName,
    projectName: usageDataProjectName,
    raisePrompt: false,
    instrumentationKey: usageData.instrumentationKeyForOfficeAddinCLITools,
    promptQuestion: "",
    usageDataLevel: sendUsageData ? usageData.UsageDataLevel.on : usageData.UsageDataLevel.off,
    method: usageData.UsageDataReportingMethod.applicationInsights,
    isForTesting: false
}
