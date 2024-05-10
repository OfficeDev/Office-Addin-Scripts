#!/usr/bin/env node

// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as commander from "commander";
import { logErrorMessage } from "office-addin-usage-data";
import * as commands from "./commands";

/* global process */

commander.name("office-addin-sso");
commander.version(process.env.npm_package_version || "(version not available)");

commander.command("configure <manifest-path> [secret-ttl]").action(commands.configureSSO);

commander.command("start <manifest-path>").action(commands.startSSOService);

// if the command is not known, display an error
commander.on("command:*", () => {
  logErrorMessage(`The command syntax is not valid.\n`);
  process.exitCode = 1;
  commander.help();
});

if (process.argv.length > 2) {
  commander.parse(process.argv);
} else {
  commander.help();
}
