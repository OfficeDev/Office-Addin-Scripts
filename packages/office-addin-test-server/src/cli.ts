// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as commander from "commander";
import { logErrorMessage } from "office-addin-cli";
import * as commands from "./commands";
import * as defaults from "./defaults";

commander.name("office-addin-test-server");
commander.version(process.env.npm_package_version || "(version not available)");

commander
    .command("start")
    .option(`--https [port number]", "Port number must be between 0 - 65535. Default: ${defaults.httpsPort}`)
    .option(`--http [port number]", "Port number must be between 0 - 65535. Default: ${defaults.httpPort}`)
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
