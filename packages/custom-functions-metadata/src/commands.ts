
// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as commander from "commander";
import { logErrorMessage } from "office-addin-cli";
import * as generateMetadata from "./generate";

export async function generate(inputFile: string, outputFile: string) {
  try {
      if (!inputFile) {
        throw new Error("You need to provide the path to the source file for custom functions.");
      }
      if (!outputFile) {
        throw new Error("You need to provide the path to the output file for the custom functions metadata.");
      }
      const results = await generateMetadata.generate(inputFile, outputFile);
      if (results.errors.length > 0) {
        console.error("Errors found:" );
        results.errors.forEach((err) => console.log(err));
      }
  } catch (err) {
    logErrorMessage(err);
  }
}
