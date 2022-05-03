#!/usr/bin/env node

// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as commander from "commander";
import { logErrorMessage } from "office-addin-usage-data";
import * as commands from "./commands";

/* global process */

commander.name("custom-functions-metadata");
commander.version(process.env.npm_package_version || "(version not available)");

commander
  .command("generate <source-file> [output-file]")
  .option("--allow-error-for-data-type-any", "Allow a custom function to process errors as input values.")
  .description("Generate the metadata for the custom functions from the source code.")
  .action(commands.generate);

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
