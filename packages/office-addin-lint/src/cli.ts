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
    .command("lint")
    .option("--files <files>", `Path to files to perform lint check. Default: ${defaults.lintFiles}`)
    .description(`Perform lint and prettier check`)
    .action(commands.lint);

commander
    .command("lint:fix")
    .option("--files <files>", `Path to files to perform lint fix. Default: ${defaults.lintFiles}`)
    .description(`Fix lint and prettier errors.`)
    .action(commands.lintFix);

commander
    .command("prettier")
    .option("--files <files>", `Path to files to perform prettier. Default: ${defaults.lintFiles}`)
    .description(`Fix all prettier issues.`)
    .action(commands.prettier);

// if the command is not known, display an error
commander.on("command:*", function() {
    logErrorMessage(`The command syntax is not valid.\n`);
    process.exitCode = 1;
    commander.help();
});

if (process.argv.length > 2) {
    commander.parse(process.argv);
} else {
    commander.help();
}
