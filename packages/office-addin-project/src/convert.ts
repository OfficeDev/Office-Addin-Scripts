// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as fs from "fs";
import * as inquirer from "inquirer";

import { exec } from "child_process";
import { ExpectedError } from "office-addin-usage-data";

/* global console */

export async function convertProject(manifestPath: string = "./manifest.xml") {
  if (manifestPath.endsWith(".json")) {
    throw new ExpectedError(
      `The convert command only works on xml manifest based projects`
    );
  }

  if (!fs.existsSync(manifestPath)) {
    throw new ExpectedError(
      `The manifest file '${manifestPath}' does not exist`
    );
  }

  const shouldContinue = await asksForUserConfirmation();
  if (!shouldContinue) {
    return;
  }
  updatePackages();
}

async function asksForUserConfirmation(): Promise<boolean> {
  const question = {
    message: `This command will convert your XML manifest to a JSON manifest and upgrade your project dependencies to make it compatible with the new project.\nWould you like to continue?`,
    name: "didUserConfirm",
    type: "confirm",
  };
  const answers = await inquirer.prompt([question]);
  return (answers as any).didUserConfirm;
}

function updatePackages(): void {
  // Contains name of the package and minimum version
  const depedentPackages: string[] = [
    "@microsoft/teamsfx",
    "office-addin-debugging",
    "office-addin-manifest",
  ];
  let command: string = "npm install";
  let messageToBePrinted: string = "Installing latest versions of";

  for (let i = 0; i < depedentPackages.length; i++) {
    const depedentPackage = depedentPackages[i];
    command += ` ${depedentPackage}@latest`;
    messageToBePrinted += ` ${depedentPackage}`;

    if (i === depedentPackages.length - 2) {
      messageToBePrinted += " and";
    } else {
      messageToBePrinted += ",";
    }
  }

  command += ` --save-dev`;
  console.log(messageToBePrinted.slice(0, -1));
  exec(command);
}
