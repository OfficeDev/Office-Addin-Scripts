// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as usageData from "office-addin-usage-data";

export const lintFiles = "src/**/*.{ts?(x),js?(x)}";

export enum ESLintExitCode {
  NoLintErrors = 0,
  HasLintError = 1,
  CommandFailed = 2,
}

export enum PrettierExitCode {
  NoFormattingProblems = 0,
  HasFormattingProblem = 1,
  CommandFailed = 2,
}

// Usage data defaults
export const usageDataObject: usageData.OfficeAddinUsageData = new usageData.OfficeAddinUsageData({
  projectName: "office-addin-lint",
  instrumentationKey: usageData.instrumentationKeyForOfficeAddinCLITools,
  raisePrompt: false,
});
