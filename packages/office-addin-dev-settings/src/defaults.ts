// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import {
  instrumentationKeyForOfficeAddinCLITools,
  OfficeAddinUsageData,
} from "office-addin-usage-data";

// Usage data defaults
export const usageDataObject: OfficeAddinUsageData = new OfficeAddinUsageData({
  projectName: "office-addin-dev-settings",
  instrumentationKey: instrumentationKeyForOfficeAddinCLITools,
  raisePrompt: false,
});
