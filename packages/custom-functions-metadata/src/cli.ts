#!/usr/bin/env node

// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import { Command } from "commander";
import { logErrorMessage } from "office-addin-usage-data";
import * as commands from "./commands";

/* global process */
const program = new Command();

program.name("custom-functions-metadata").version(process.env.npm_package_version || "(version not available)");

program
  .command("generate")
  .argument(
    "<source-files>",
    "The path to the source file (or comma separated list of files) for custom functions.",
    commaSeparatedList
  )
  .argument("[output-file]", "The path to the output file for the metadata.")
  .option("--allow-custom-data-for-data-type-any", "Allow custom functions to return FormattedNumberCellValues.")
  .description("Generate the metadata for the custom functions from the source code files.")
  .action(commands.generate);

// if the command is not known, display an error
program.on("command:*", function () {
  logErrorMessage(`The command syntax is not valid.\n`);
  process.exitCode = 1;
  program.help();
});

if (process.argv.length > 2) {
  program.parse(process.argv);
} else {
  program.help();
}

// eslint-disable-next-line @typescript-eslint/no-unused-vars
function commaSeparatedList(value: string, dummyPrevious: string): string | string[] {
  value = value.trim();
  if (value.indexOf(",") < 0) {
    return value;
  } else {
    return value.split(",");
  }
}
