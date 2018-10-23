#!/usr/bin/env node

// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as commander from "commander";
import * as commands from "./commands";

if (process.argv[1].endsWith("\\manifest.js")) {
  commander
    .command("info <path>")
    .action(commands.info);

  commander
    .command("modify <path>")
    .option("-g,--guid [guid]", "Change the guid. Specify \"random\" for a random guid.")
    .option("-d, --displayName [name]", "Display name for the add-in.")
    .action(commands.modify);

  commander.parse(process.argv);
}
