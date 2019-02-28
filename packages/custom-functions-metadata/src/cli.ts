#!/usr/bin/env node

// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as commander from "commander";
import * as commands from "./commands";

commander
  .command("generate <source-file> <metadata-file>")
  .description("Generate the metadata for the custom functions from the source code.")
  .action(commands.generate);

commander.parse(process.argv);
