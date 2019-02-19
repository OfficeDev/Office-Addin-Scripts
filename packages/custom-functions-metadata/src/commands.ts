
// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as commander from "commander";
import * as generateMetadata from "./generate";

export async function generate(inputFile: string, outputFile: string) {
  try {
      if (!inputFile) {
        throw new Error("You need to provide the path to the source file for custom functions.");
      }
      if (!outputFile) {
        throw new Error("You need to provide the path to the output file for the custom functions metadata.");
      }
      const errors = await generateMetadata.generate(inputFile, outputFile);
      if (errors.length > 0) {
        console.log("Errors found:" );
        errors.forEach((err) => console.log(err));
      }
  } catch (err) {
    logErrorMessage(err);
  }
}

function logErrorMessage(err: any) {
  console.error(`Error: ${err instanceof Error ? err.message : err}`);
}
