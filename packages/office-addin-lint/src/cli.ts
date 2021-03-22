#!/usr/bin/env node

// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as commander from "commander";
import { logErrorMessage } from "office-addin-cli";
import * as commands from "./commands";
import * as defaults from "./defaults";

commander.name("office-addin-lint");
commander.version(process.env.npm_package_version || "(version not available)");

commander
  .command("check")
  .option("--files <files>", `Specifies the source files to check. Default: ${defaults.lintFiles}`)
  .description(`Check source files against lint rules.`)
  .action(commands.lint);

commander
  .command("fix")
  .option("--files <files>", `Specifies the source files to fix. Default: ${defaults.lintFiles}`)
  .description(`Apply fixes to source based on lint rules.`)
  .action(commands.lintFix);

commander
  .command("prettier")
  .option("--files <files>", `Specifies which files to use. Default: ${defaults.lintFiles}`)
  .description(`Make the source prettier.`)
  .action(commands.prettier);

// if the command is not known, display an error
commander.on("command:*", function () {
  logErrorMessage(`The command syntax is not valid.\n`);
  process.exitCode = 1;
  commander.help();
});

if (process.argv.length > 2) {
  commander.parse(process.argv);
} else {
  commander.help();
}
