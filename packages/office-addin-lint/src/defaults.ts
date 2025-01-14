// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import { instrumentationKeyForOfficeAddinCLITools, OfficeAddinUsageData } from "office-addin-usage-data";

export const lintFiles = "src/**/*.{ts,tsx,js,jsx}";

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
export const usageDataObject: OfficeAddinUsageData = new OfficeAddinUsageData({
  projectName: "office-addin-lint",
  instrumentationKey: instrumentationKeyForOfficeAddinCLITools,
  raisePrompt: false,
});
