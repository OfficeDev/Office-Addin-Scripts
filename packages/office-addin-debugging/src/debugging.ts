#!/usr/bin/env node

// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as commander from "commander";
import * as commands from "./commands";

export * from "./start";
export * from "./stop";

if (process.argv[1].endsWith("\\debugging.js")) {
    commander
        .command("start <manifestPath>")
        .option("--debug-method <method>", "The debug method to use.")
        .option("--dev-server <command>", "Run the dev server.")
        .option("--dev-server-url <url>", "Verify dev server is running using this url.")
        .option("--no-debug", "Start without debugging.")
        .option("--packager <command>", "Run the packager.")
        .option("--packager-host <host>")
        .option("--packager-port <port>")
        .option("--sideload <command>")
        .option("--source-bundle-url-host <host>")
        .option("--source-bundle-url-port <port>")
        .option("--source-bundle-url-path <path>")
        .option("--source-bundle-url-extension <extension>")
        .action(commands.start);

    commander
        .command("stop <manifestPath>")
        .option("--unload <command>")
        .action(commands.stop);

    commander.parse(process.argv);
}
