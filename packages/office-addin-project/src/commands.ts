// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as commander from "commander";
import * as inquirer from "inquirer";
import { logErrorMessage } from "office-addin-usage-data";
import { usageDataObject } from "./defaults";
import { convertProject } from "./convert";

export async function convert(command: commander.Command) {
  try {
    const manifestPath: string = command.manifest ?? "./manifest.xml";
    const backupPath: string = command.backup ?? "./backup.zip";
    const projectPath: string = command.project ?? "";
    const devPreview: boolean = command.preview ?? false;
    const shouldContinue = command.confirm ?? await asksForUserConfirmation();

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
  const question = {
    message: `This command will convert your current xml manifest to a json manifest and then proceed to upgrade your project dependencies to ensure compatibility with the new project structure.\nHowever, in order for this newly updated project to function correctly you must be using a compatible version of Outlook.\nWould you like to continue?`,
    name: "didUserConfirm",
    type: "confirm",
  };
  const answers = await inquirer.prompt([question]);
  return (answers as any).didUserConfirm;
}
