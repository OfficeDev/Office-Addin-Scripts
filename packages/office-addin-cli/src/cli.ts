#!/usr/bin/env node

// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as commander from "commander";
import { logErrorMessage } from "./log";
import * as commands from "./commands";

/* global process */

commander.name("office-addin-cli");
commander.version(process.env.npm_package_version || "(version not available)");

// if the command is not known, display an error
commander.on("command:*", function () {
  logErrorMessage(`The command syntax is not valid.\n`);
  process.exitCode = 1;
  commander.help();
});

commander.command("convert <manifest-path>").action(commands.convert);

if (process.argv.length > 2) {
  commander.parse(process.argv);
} else {
  commander.help();
}
