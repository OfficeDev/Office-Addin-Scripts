// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as commander from "commander";
import {logErrorMessage, parseNumber} from "office-addin-cli";
import * as defaults from "./defaults";
import * as prettierLint from "./lint";

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
    const pathToFiles: string = getPathToFiles(command);
    prettierLint.performLintCheck(pathToFiles);
}

export async function lintFix(command: commander.Command) {
    const pathToFiles: string = getPathToFiles(command);
    prettierLint.performLintFix(pathToFiles);
}

export async function prettier(command: commander.Command) {
    const pathToFiles: string = getPathToFiles(command);
    prettierLint.makeFilesPrettier(pathToFiles);
}
