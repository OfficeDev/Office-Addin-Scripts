// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import { logErrorMessage } from "./log";
import { usageDataObject } from "./defaults";

export async function upgrade(manifestPath: string) {
  try {
    // Convert manifest .xml to manifest.json
    // Do any more needed modifications on the code?

    usageDataObject.reportSuccess("upgrade");
    throw new Error("Upgrade function is not ready yet.");
  } catch (err: any) {
    usageDataObject.reportException("upgrade", err);
    logErrorMessage(err);
  }
}
