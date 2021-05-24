// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as path from "path";
import { usageDataObject, ESLintExitCode, PrettierExitCode } from "./defaults";

const eslintPath = require.resolve("eslint");
const prettierPath = require.resolve("prettier");
const eslintDir = path.parse(eslintPath).dir;
const eslintFilePath = path.resolve(eslintDir, "../bin/eslint.js");
const prettierFilePath = path.resolve(prettierPath, "../bin-prettier.js");
const eslintConfigPath = path.resolve(__dirname, "../config/.eslintrc.json");
const eslintTestConfigPath = path.resolve(__dirname, "../config/.eslintrc.test.json");

function execCommand(command: string) {
  const execSync = require("child_process").execSync;
  const child = execSync(command, { stdio: "inherit" });
}

function normalizeFilePath(filePath: string): string {
  return filePath.replace(/ /g, "\\ "); // Converting space to '\\'
}

function getEsLintBaseCommand(useTestConfig: boolean): string {
  const configFilePath = useTestConfig ? eslintTestConfigPath : eslintConfigPath
  const eslintBaseCommand: string = `node ${eslintFilePath} -c ${configFilePath} --resolve-plugins-relative-to ${__dirname}`;
  return eslintBaseCommand;
}

export function getLintCheckCommand(files: string, useTestConfig: boolean = false): string {
  const eslintCommand: string = `${getEsLintBaseCommand(useTestConfig)} ${normalizeFilePath(files)}`;
  return eslintCommand;
}

export function performLintCheck(files: string, useTestConfig: boolean) {
  try {
    const command = getLintCheckCommand(files, useTestConfig);
    execCommand(command);
    usageDataObject.reportSuccess("performLintCheck()", { exitCode: ESLintExitCode.NoLintErrors });
  } catch (err) {
    if (err.status && err.status == ESLintExitCode.HasLintError) {
      usageDataObject.reportExpectedException("performLintCheck()", err, { exitCode: ESLintExitCode.HasLintError });
    } else {
      usageDataObject.reportException("performLintCheck()", err);
    }
    throw err;
  }
}

export function getLintFixCommand(files: string, useTestConfig: boolean = false): string {
  const eslintCommand: string = `${getEsLintBaseCommand(useTestConfig)} --fix ${normalizeFilePath(files)}`;
  return eslintCommand;
}

export function performLintFix(files: string, useTestConfig: boolean) {
  try {
    const command = getLintFixCommand(files, useTestConfig);
    execCommand(command);
    usageDataObject.reportSuccess("performLintFix()", { exitCode: ESLintExitCode.NoLintErrors });
  } catch (err) {
    if (err.status && err.status == ESLintExitCode.HasLintError) {
      usageDataObject.reportExpectedException("performLintFix()", err, { exitCode: ESLintExitCode.HasLintError });
    } else {
      usageDataObject.reportException("performLintFix()", err);
    }
    throw err;
  }
}

export function getPrettierCommand(files: string): string {
  const prettierFixCommand: string = `node ${prettierFilePath} --parser typescript --write ${normalizeFilePath(files)}`;
  return prettierFixCommand;
}

export function makeFilesPrettier(files: string) {
  try {
    const command = getPrettierCommand(files);
    execCommand(command);
    usageDataObject.reportSuccess("makeFilesPrettier()", { exitCode: PrettierExitCode.NoFormattingProblems });
  } catch (err) {
    if (err.status && err.status == PrettierExitCode.HasFormattingProblem) {
      usageDataObject.reportExpectedException("makeFilesPrettier()", err, {
        exitCode: PrettierExitCode.HasFormattingProblem,
      });
    } else {
      usageDataObject.reportException("makeFilesPrettier()", err);
    }
    throw err;
  }
}
