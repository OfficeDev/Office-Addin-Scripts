// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import {logErrorMessage} from "office-addin-cli";
import * as defaults from "./defaults";

function execCommand(command: string) {
    const execSync = require('child_process').execSync;
    const child = execSync(command, { stdio: 'inherit' });
}

export function lintCheckCommand(files: string): string {
  const eslintCommand: string = "eslint " + files;
  return eslintCommand;
}

export function lintCheck(files: string) {
  console.log("Lint check for files: " + files);
  const command = lintCheckCommand(files);
  execCommand(command);
}

export function lintFixCommand(files: string): string {
  const eslintCommand: string = "eslint --fix " + files;
  return eslintCommand;
}

export function lintFix(files: string) {
    console.log("Fixing lint issues for files: " + files);
    const command = lintFixCommand(files);
    execCommand(command);
}

export function prettierCommand(files: string): string {
  const prettierFixCommand: string = "prettier --parser typescript --write " + files;
  return prettierFixCommand;
}

export function prettier(files: string) {
    console.log("Fixing prettier issues for files: " + files);
    const command = prettierCommand(files);
    execCommand(command);
}
