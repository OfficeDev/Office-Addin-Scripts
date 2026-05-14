// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import {
  connectionStringForOfficeAddinCLITools,
  OfficeAddinUsageData,
} from "office-addin-usage-data";

// Usage data defaults
export const usageDataObject: OfficeAddinUsageData = new OfficeAddinUsageData({
  projectName: "office-addin-debugging",
  connectionString: connectionStringForOfficeAddinCLITools,
  raisePrompt: false,
});
