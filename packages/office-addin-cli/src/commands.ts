// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as commander from "commander";
import { logErrorMessage } from "./log";
import { usageDataObject } from "./defaults";

/* global console, process */

export async function upgrade(
  manifestPath: string,
  command: commander.Command /* eslint-disable-line @typescript-eslint/no-unused-vars */
) {
  try {
    usageDataObject.reportSuccess("upgrade");
  } catch (err: any) {
    usageDataObject.reportException("upgrade", err);
    logErrorMessage(err);
  }
}
