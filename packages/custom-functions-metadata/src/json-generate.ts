#!/usr/bin/env node

// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as commander from "commander";
import * as commands from "./commands";
import * as metadata from "./custom-functions-metadata";

export async function generate(inputFile:string, outputFile:string) {
    await metadata.generate(inputFile,outputFile);
}

if (process.argv[1].endsWith("\\json-generate.js")) {
    commander
      .command("generate <jsourceFile> <metadataFile>")
      .description("Generate the metadata for the custom functions from the source code.")
      .action(commands.generate);
  
    commander.parse(process.argv);
  }