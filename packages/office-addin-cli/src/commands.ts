// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as commander from "commander";
import { logErrorMessage } from "office-addin-usage-data";
import { usageDataObject } from "./defaults";
import { convertProject } from "./convert";

export async function convert(command: commander.Command) {
  try {
    const manifestPath: string = command.manifest ?? "./manifest.xml";
    const packageJsonPath: string = command.packageJson ?? "./package.json";

    await convertProject(manifestPath, packageJsonPath);
    usageDataObject.reportSuccess("convert");
    throw new Error("Upgrade function is not ready yet.");
  } catch (err: any) {
    usageDataObject.reportException("convert", err);
    logErrorMessage(err);
  }
}
