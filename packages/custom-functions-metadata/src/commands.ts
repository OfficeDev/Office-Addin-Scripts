// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as commander from "commander";
import { writeFileSync } from "fs";
import { logErrorMessage } from "office-addin-cli";
import { generateCustomFunctionsMetadata }  from "./generate";

export async function generate(inputPath: string, outputPath: string) {
  try {
    if (!inputPath) {
      throw new Error("You need to provide the path to the source file for custom functions.");
    }
    const results = await generateCustomFunctionsMetadata(inputPath);
    if (results.errors.length > 0) {
      console.error("Errors found:");
      results.errors.forEach(err => console.log(err));
    } else {
      if (outputPath) {
        try {
          writeFileSync(outputPath, results.metadataJson);
        } catch (err) {
          throw new Error(`Cannot write to file: ${outputPath}.`);
        }
      } else {
        console.log(results.metadataJson);
      }
    }
  } catch (err) {
    logErrorMessage(err);
  }
}
