#!/usr/bin/env node

// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as commander from "commander";
import * as commands from "./commands";
import * as cfm from "./custom-functions-metadata";

export async function generate(inputfile:string, outputfile:string) {
    cfm.generate(inputfile,outputfile);
}

if (process.argv[1].endsWith("\\json-generate.js")) {
    commander
      .command("generate <inputfile> <outputfile>")
      .description("Generate the custom functions json.")
      .action(commands.generate);
  
    commander.parse(process.argv);
  }