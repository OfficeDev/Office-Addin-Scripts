// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import {logErrorMessage} from "office-addin-cli";
import * as defaults from "./defaults";

function execCommand(command: string) {
    const execSync = require('child_process').execSync;
    const child = execSync(command, { stdio: 'inherit' });
}

export function getLintCheckCommand(files: string): string {
  const eslintCommand: string = "eslint " + files;
  return eslintCommand;
}

export function performLintCheck(files: string) {
  const command = getLintCheckCommand(files);
  execCommand(command);
}

export function getLintFixCommand(files: string): string {
  const eslintCommand: string = "eslint --fix " + files;
  return eslintCommand;
}

export function performLintFix(files: string) {
    const command = getLintFixCommand(files);
    execCommand(command);
}

export function getPrettierCommand(files: string): string {
  const prettierFixCommand: string = "prettier --parser typescript --write " + files;
  return prettierFixCommand;
}

export function makeFilesPrettier(files: string) {
    const command = getPrettierCommand(files);
    execCommand(command);
}
