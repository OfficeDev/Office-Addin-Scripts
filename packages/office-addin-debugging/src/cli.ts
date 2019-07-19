#!/usr/bin/env node

// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as commander from "commander";
import { logErrorMessage } from "office-addin-cli";
import * as commands from "./commands";

commander.name("office-addin-debugging");
commander.version(process.env.npm_package_version || "(version not available)");

commander
    .command("start <manifest-path> [app-type]")
    .option("--app <app>", "Specify which Office app to use.")
    .option("--debug-method <method>", "The debug method to use.")
    .option("--dev-server <command>", "Run the dev server.")
    .option("--dev-server-port <port>", "Verify the dev server is running using this port.")
    .option("--no-debug", "Start without debugging.")
    .option("--no-live-reload", "Do not enable live-reload.")
    .option("--packager <command>", "Run the packager.")
    .option("--packager-host <host>")
    .option("--packager-port <port>")
    .option("--prod", "Use the production bundle")
    .option("--sideload <command>")
    .option("--source-bundle-url-host <host>")
    .option("--source-bundle-url-port <port>")
    .option("--source-bundle-url-path <path>")
    .option("--source-bundle-url-extension <extension>")
    .action(commands.start);

commander
    .command("stop <manifest-path> [app-type]")
    .option("--prod", "Use the production bundle")
    .option("--unload <command>")
    .action(commands.stop);

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
