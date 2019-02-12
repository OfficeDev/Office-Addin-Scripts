#!/usr/bin/env node

// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as commander from "commander";
import * as path from "path";
import * as commands from "./commands";

export * from "./generate";

// if this package is being run from the command line
// (related to "main" and "bin" in package.json)
if (process.argv[1].endsWith(path.join("lib", "custom-functions-metadata.js"))
  || process.argv[1].endsWith(path.join(".bin", "custom-functions-metadata"))) {
  commander
    .command("generate <jsourceFile> <metadataFile>")
    .description("Generate the metadata for the custom functions from the source code.")
    .action(commands.generate);

  commander.parse(process.argv);
}
