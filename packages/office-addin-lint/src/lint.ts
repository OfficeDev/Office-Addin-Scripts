// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as path from "path";
import * as usageDataHelper from "./usagedata-helper"

const esLintPath = require.resolve("eslint");
const prettierPath = require.resolve("prettier");
const esLintDir = path.parse(esLintPath).dir;
const esLintFilePath = path.resolve(esLintDir, "../bin/eslint.js");
const prettierFilePath = path.resolve(prettierPath, "../bin-prettier.js");
const esLintConfigPath = path.resolve(__dirname, "../config/.eslintrc.json");

function execCommand(command: string) {
  try {
    const execSync = require("child_process").execSync;
    const child = execSync(command, { stdio: "inherit" });
  } catch (err) {
    throw err;
  } 
}

function getEsLintBaseCommand(): string {
  const eslintBaseCommand: string = `node ${esLintFilePath} -c ${esLintConfigPath} --resolve-plugins-relative-to ${__dirname}`;
  return eslintBaseCommand;
}

export function getLintCheckCommand(files: string): string {
  const eslintCommand: string = `${getEsLintBaseCommand()} ${files}`;
  return eslintCommand;
}

export function performLintCheck(files: string) {
  try {
    const command = getLintCheckCommand(files);
    execCommand(command);
    usageDataHelper.sendUsageDataSuccessEvent("check");
  } catch (err) {
    usageDataHelper.sendUsageDataException("check", `${err}`);
    throw err;
  }
}

export function getLintFixCommand(files: string): string {
  const eslintCommand: string = `${getEsLintBaseCommand()} --fix ${files}`;
  return eslintCommand;
}

export function performLintFix(files: string) {
  try{
    const command = getLintFixCommand(files);
    execCommand(command);
    usageDataHelper.sendUsageDataSuccessEvent("fix");
  } catch (err) {
    usageDataHelper.sendUsageDataException("fix", `${err}`);
    throw err;
  }
}

export function getPrettierCommand(files: string): string {
  const prettierFixCommand: string = `node ${prettierFilePath} --parser typescript --write ${files}`;
  return prettierFixCommand;
}

export function makeFilesPrettier(files: string) {
  try{
    const command = getPrettierCommand(files);
    execCommand(command);
    usageDataHelper.sendUsageDataSuccessEvent("prettier");
  } catch (err) {
    usageDataHelper.sendUsageDataException("prettier", `${err}`);
    throw err;
  }
}
