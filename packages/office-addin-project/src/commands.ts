// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import { OptionValues } from "commander";
import inquirer from "inquirer";
import { logErrorMessage } from "office-addin-usage-data";
import { usageDataObject } from "./defaults";
import { convertProject } from "./convert";

export async function convert(options: OptionValues) {
  try {
    const manifestPath: string = options.manifest ?? "./manifest.xml";
    const backupPath: string = options.backup ?? "./backup.zip";
    const projectPath: string = options.project ?? "";
    const devPreview: boolean = options.preview ?? false;
    const shouldContinue = options.confirm ?? (await asksForUserConfirmation());

    if (shouldContinue) {
      await convertProject(manifestPath, backupPath, projectPath, devPreview);
      usageDataObject.reportSuccess("convert", { result: "Project converted" });
    } else {
      usageDataObject.reportSuccess("convert", {
        result: "Conversion cancelled",
      });
    }
  } catch (err: any) {
    usageDataObject.reportException("convert", err);
    logErrorMessage(err);
  }
}

async function asksForUserConfirmation(): Promise<boolean> {
  const answers = await inquirer.prompt({
    message: `This command will convert your current xml manifest to a json manifest and then proceed to upgrade your project dependencies to ensure compatibility with the new project structure.\nHowever, in order for this newly updated project to function correctly you must be using a compatible version of Office applications.\nWould you like to continue?`,
    name: "didUserConfirm",
    type: "confirm",});
  return (answers as any).didUserConfirm;
}
