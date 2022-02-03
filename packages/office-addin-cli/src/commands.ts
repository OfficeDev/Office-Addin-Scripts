// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import { logErrorMessage } from "office-addin-usage-data";
import { usageDataObject } from "./defaults";

export async function convert(manifestPath: string) {
  try {
    // Convert manifest .xml to manifest.json
    // Do any more needed modifications on the code?
    usageDataObject.reportSuccess("convert");
    throw new Error("Upgrade function is not ready yet.");
  } catch (err: any) {
    usageDataObject.reportException("convert", err);
    logErrorMessage(err);
  }
}
