// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as commander from "commander";
import { logErrorMessage } from "office-addin-cli";
import * as defaults from "./defaults";
import { makeFilesPrettier, performLintCheck, performLintFix } from "./lint";

/**
 * Determines path to files to run lint against. Priority order follows:
 * 1 --files option passed through cli
 * 2 lint_files config property in package.json
 * 3 default location
 * @param command command options which can contain files
 */
function getPathToFiles(command: commander.Command): string {
  const pathToFiles: any = command.files
    ? command.files
    : process.env.npm_package_config_lint_files;
  return pathToFiles ? pathToFiles : defaults.lintFiles;
}

export async function lint(command: commander.Command) {
  try {
    const pathToFiles: string = getPathToFiles(command);
    performLintCheck(pathToFiles);
  } catch (err) {
    // no need to display an error since there will already be error output;
    // just return a non-zero exit code
    process.exitCode = 1;
  }
}

export async function lintFix(command: commander.Command) {
  try {
    const pathToFiles: string = getPathToFiles(command);
    performLintFix(pathToFiles);
  } catch (err) {
    // no need to display an error since there will already be error output;
    // just return a non-zero exit code
    process.exitCode = 1;
  }
}

export async function prettier(command: commander.Command) {
  try {
    const pathToFiles: string = getPathToFiles(command);
    makeFilesPrettier(pathToFiles);
  } catch (err) {
    logErrorMessage(`Unable to make code prettier.\n${err}`);
    process.exitCode = 1;
  }
}
