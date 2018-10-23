
// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as commander from "commander";
import * as metadata from "./custom-functions-metadata";

export async function generate(inputFile: string, outputFile: string) {
  try {
      if (!inputFile) {
        throw new Error("You need to provide the path to the source file for custom functions.");
      }
      if (!outputFile) {
        throw new Error("You need to provide the path to the output file for the custom functions metadata.");
      }
      console.log("Begin json generation")
      await metadata.generate(inputFile,outputFile);
  }
  catch (err){
    console.log('Error: ' + err);
  }
}
