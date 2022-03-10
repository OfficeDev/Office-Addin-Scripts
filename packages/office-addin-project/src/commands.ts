// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as commander from "commander";
import { logErrorMessage } from "office-addin-usage-data";
import { usageDataObject } from "./defaults";
import { convertProject } from "./convert";

export async function convert(command: commander.Command) {
  try {
    const manifestPath: string = command.manifest ?? "./manifest.xml";

    await convertProject(manifestPath);
    throw new Error("Upgrade function is not ready yet.");
    // usageDataObject.reportSuccess("convert");
  } catch (err: any) {
    usageDataObject.reportException("convert", err);
    logErrorMessage(err);
  }
}
