// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as usageData from "office-addin-usage-data";

// Usage data defaults
export const usageDataObject: usageData.OfficeAddinUsageData =
  new usageData.OfficeAddinUsageData({
    projectName: "eslint-plugin-office-addins",
    instrumentationKey: usageData.instrumentationKeyForOfficeAddinCLITools,
    raisePrompt: false,
  });
