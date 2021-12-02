// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as commander from "commander";
import { logErrorMessage } from "office-addin-cli";
import * as defaults from "./defaults";
import { makeFilesPrettier, performLintCheck, performLintFix } from "./lint";
import { usageDataObject } from "./defaults";

/* global process */

/**
 * Determines path to files to run lint against. Priority order follows:
 * 1 --files option passed through cli
 * 2 lint_files config property in package.json
 * 3 default location
 * @param command command options which can contain files
 */
function getPathToFiles(command: commander.Command): string {
  const pathToFiles: any = command.files ? command.files : process.env.npm_package_config_lint_files;
  return pathToFiles ? pathToFiles : defaults.lintFiles;
}

export async function lint(command: commander.Command) {
  try {
    const pathToFiles: string = getPathToFiles(command);
    const useTestConfig: boolean = command.test;
    await performLintCheck(pathToFiles, useTestConfig);
    usageDataObject.reportSuccess("lint");
  } catch (err: any) {
    if (typeof err.status == "number") {
      process.exitCode = err.status;
    } else {
      process.exitCode = defaults.ESLintExitCode.CommandFailed;
      usageDataObject.reportException("lint", err);
      logErrorMessage(err);
    }
  }
}

export async function lintFix(command: commander.Command) {
  try {
    const pathToFiles: string = getPathToFiles(command);
    const useTestConfig: boolean = command.test;
    await performLintFix(pathToFiles, useTestConfig);
    usageDataObject.reportSuccess("lintFix");
  } catch (err: any) {
    if (typeof err.status == "number") {
      process.exitCode = err.status;
    } else {
      process.exitCode = defaults.ESLintExitCode.CommandFailed;
      usageDataObject.reportException("lintFix", err);
      logErrorMessage(err);
    }
  }
}

export async function prettier(command: commander.Command) {
  try {
    const pathToFiles: string = getPathToFiles(command);
    await makeFilesPrettier(pathToFiles);
    usageDataObject.reportSuccess("prettier");
  } catch (err: any) {
    if (typeof err.status == "number") {
      process.exitCode = err.status;
    } else {
      process.exitCode = defaults.PrettierExitCode.CommandFailed;
      usageDataObject.reportException("prettier", err);
      logErrorMessage(err);
    }
  }
}
