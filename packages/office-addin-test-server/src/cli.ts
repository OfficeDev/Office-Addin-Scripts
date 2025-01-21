// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import { Command } from "commander";
import { logErrorMessage } from "office-addin-usage-data";
import * as commands from "./commands";
import { defaultPort } from "./testServer";

const commander = new Command();

commander.name("office-addin-test-server");
commander.version(process.env.npm_package_version || "(version not available)");

commander
  .command("start")
  .option(
    `-p --port [port number]", "Port number must be between 0 - 65535. Default: ${defaultPort}`
  )
  .action(commands.start);

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
