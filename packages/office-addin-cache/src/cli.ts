#!/usr/bin/env node

// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import { Command } from "commander";
import { logErrorMessage } from "office-addin-usage-data";
import * as commands from "./commands";

/* global process */

const commander = new Command();

commander.name("office-addin-cache");
commander.version(process.env.npm_package_version || "(version not available)");

commander
  .command("clear")
  .description("Empty the Office Add-in cache folders.")
  .option("--force-close", "Close all running Microsoft Office applications without prompting.")
  .option("--verbose", "Report the path of each folder being emptied.")
  .action(commands.clear);

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
