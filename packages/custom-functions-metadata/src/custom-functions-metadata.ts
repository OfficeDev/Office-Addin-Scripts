#!/usr/bin/env node

// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as commander from "commander";
import * as commands from "./commands";

if (process.argv[1].endsWith("\\custom-functions-metadata.js")) {
  commander
    .command("generate")
    .description("Generate metadata for custom functions.")
    .action(commands.generate);

  commander.parse(process.argv);
}
