// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import fs from "fs";
import path from "path";
import { ESLintExitCode, PrettierExitCode } from "./defaults";

/* global process require __dirname */

const eslintPath = require.resolve("eslint");
const prettierPath = require.resolve("prettier");
const eslintDir = path.parse(eslintPath).dir;
const eslintFilePath = path.resolve(eslintDir, "../bin/eslint.js");
const prettierFilePath = path.resolve(prettierPath, "../bin/prettier.cjs");
const eslintConfigPath = path.resolve(__dirname, "../config/eslint.config.mjs");
const eslintTestConfigPath = path.resolve(__dirname, "../config/eslint.config.test.mjs");

function execCommand(command: string) {
  const execSync = require("child_process").execSync;
  execSync(command, { stdio: "inherit" });
}

function normalizeFilePath(filePath: string): string {
  return filePath.replace(/ /g, "\\ "); // Converting space to '\\'
}

function getEsLintBaseCommand(useTestConfig: boolean = false): string {
  const projLintConfig = path.resolve(process.cwd(), "eslint.config.mjs");
  const prodConfig = fs.existsSync(projLintConfig) ? projLintConfig : eslintConfigPath;
  const configFilePath = useTestConfig ? eslintTestConfigPath : prodConfig;
  const eslintBaseCommand: string = `node "${eslintFilePath}" -c "${configFilePath}"`;
  return eslintBaseCommand;
}

export function getLintCheckCommand(files: string, useTestConfig: boolean = false): string {
  const eslintCommand: string = `${getEsLintBaseCommand(useTestConfig)} "${normalizeFilePath(files)}"`;
  return eslintCommand;
}

export function performLintCheck(files: string, useTestConfig: boolean = false) {
  const command = getLintCheckCommand(files, useTestConfig);
  execCommand(command);
}

export function getLintFixCommand(files: string, useTestConfig: boolean = false): string {
  const eslintCommand: string = `${getEsLintBaseCommand(useTestConfig)} --fix ${normalizeFilePath(files)}`;
  return eslintCommand;
}

export function performLintFix(files: string, useTestConfig: boolean = false) {
  const command = getLintFixCommand(files, useTestConfig);
  execCommand(command);
}

export function getPrettierCommand(files: string): string {
  const prettierFixCommand: string = `node ${prettierFilePath} --parser typescript --write ${normalizeFilePath(files)}`;
  return prettierFixCommand;
}

export function makeFilesPrettier(files: string) {
  const command = getPrettierCommand(files);
  execCommand(command);
}
