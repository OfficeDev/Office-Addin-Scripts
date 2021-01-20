// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as usageData from "office-addin-usage-data";

export const lintFiles = "src/**/*.{ts,tsx,js,jsx}";

export enum ESLintExitCode {
  Success = 0,
  LintingError = 1,
  ToolingError = 2
}

export enum PrettierExitCode {
  Success = 0,
  LintingError = 1,
  ToolingError = 2
}

// Usage data defaults
export const usageDataObject: usageData.OfficeAddinUsageData = new usageData.OfficeAddinUsageData({
  projectName: "office-addin-lint",
  instrumentationKey: usageData.instrumentationKeyForOfficeAddinCLITools,
  raisePrompt: false
});
