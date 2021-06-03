#!/usr/bin/env node

// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as commander from "commander";
import { logErrorMessage } from "office-addin-cli";
import * as commands from "./commands";

/* global process */

commander.name("office-addin-manifest");
commander.version(process.env.npm_package_version || "(version not available)");

commander.command("info <manifest-path>").action(commands.info);

commander
  .command("modify <manifest-path>")
  .option(
    "-g,--guid [guid]",
    "Change the guid. A random guid is used if one is not provided."
  )
  .option("-d, --displayName <name>", "Change the display name.")
  .action(commands.modify);

commander.command("validate <manifest-path>").action(commands.validate);

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
