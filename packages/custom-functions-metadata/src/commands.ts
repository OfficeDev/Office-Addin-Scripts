// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import { writeFileSync } from "fs";
import { logErrorMessage } from "office-addin-usage-data";
import { generateCustomFunctionsMetadata, MetadataOptions } from "./generate";

/* global console */

export async function generate(input: string | string[], outputPath: string, options: any) {
  try {
    if (!input) {
      throw new Error("You need to provide the path to the source file for custom functions.");
    }
    const additionalOptions: MetadataOptions | null = options;
    const results = await generateCustomFunctionsMetadata(input, false, additionalOptions);
    if (results.errors.length > 0) {
      console.error("Errors found:");
      results.errors.forEach((err) => console.log(err));
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
