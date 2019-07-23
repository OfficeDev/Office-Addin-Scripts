// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as path from "path";

const esLintPath = require.resolve("eslint");
const prettierPath = require.resolve("prettier");
const config = require.resolve("office-addin-lint");
const configDir = path.parse(config).dir;
const esLintDir = path.parse(esLintPath).dir;
const esLintFilePath = path.resolve(esLintDir, "\..\\bin/eslint.js");
const prettierFilePath = path.resolve(prettierPath, "\..\\bin-prettier.js");
const esLintConfigPath = path.resolve(configDir, "\..\\src/.eslintrc.json");

function execCommand(command: string) {
    const execSync = require("child_process").execSync;
    const child = execSync(command, { stdio: "inherit" });
}

function getEsLintBaseCommand(): string {
  const eslintBaseCommand: string = "node " + esLintFilePath + " -c " + esLintConfigPath + " --resolve-plugins-relative-to " + configDir + " ";
  return eslintBaseCommand;
}

export function getLintCheckCommand(files: string): string {
  const eslintCommand: string = getEsLintBaseCommand() + files;
  return eslintCommand;
}

export function performLintCheck(files: string) {
  const command = getLintCheckCommand(files);
  execCommand(command);
}

export function getLintFixCommand(files: string): string {
  const eslintCommand: string =  getEsLintBaseCommand() + "--fix " + files;
  return eslintCommand;
}

export function performLintFix(files: string) {
    const command = getLintFixCommand(files);
    execCommand(command);
}

export function getPrettierCommand(files: string): string {
  const prettierFixCommand: string = "node " + prettierFilePath + " --parser typescript --write " + files;
  return prettierFixCommand;
}

export function makeFilesPrettier(files: string) {
    const command = getPrettierCommand(files);
    execCommand(command);
}
