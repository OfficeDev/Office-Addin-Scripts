// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import { OptionValues } from "commander";
import { logErrorMessage } from "office-addin-usage-data";
import { clearCache } from "./cache";
import { usageDataObject } from "./defaults";

/* global process */

export async function clear(options: OptionValues): Promise<void> {
  try {
    await clearCache({
      forceClose: options.forceClose ?? false,
      verbose: options.verbose ?? false,
    });
    usageDataObject.reportSuccess("clear");
  } catch (err: any) {
    process.exitCode = 1;
    usageDataObject.reportException("clear", err);
    logErrorMessage(err);
  }
}
