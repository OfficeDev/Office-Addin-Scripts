#!/usr/bin/env node

// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as commander from "commander";
import * as commands from "./commands";

commander.name("custom-functions-metadata");
commander.version(process.env.npm_package_version || "(version not available)");

commander
  .command("generate <source-file> <metadata-file>")
  .description("Generate the metadata for the custom functions from the source code.")
  .action(commands.generate);

if (process.argv.length > 2) {
    commander.parse(process.argv);
} else {
    commander.help();
}
