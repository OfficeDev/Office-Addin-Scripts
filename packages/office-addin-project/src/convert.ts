// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as AdmZip from "adm-zip";
import * as fs from "fs";
import * as fsExtra from "fs-extra";
import * as inquirer from "inquirer";
import * as path from "path";
import * as util from "util";
import { exec } from "child_process";
import { ExpectedError } from "office-addin-usage-data";

/* global console */

const skipBackup: string[] = [ "node_modules" ]

export async function convertProject(
  manifestPath: string = "./manifest.xml",
  backupPath: string = "./backup.zip"
) {
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
  await backupProject(backupPath);
  updatePackages();
  await updateManifestXmlReferences();
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

async function backupProject(backupPath: string) {
  const zip: AdmZip = new AdmZip();
  const outputPath: string = path.resolve(backupPath);
  const rootDir: string = path.resolve(); 

  const files: string[] = fs.readdirSync(rootDir);
  files.forEach((entry) => {
    const fullPath = path.resolve(entry)
    const entryStats = fs.lstatSync(fullPath);

    if (skipBackup.includes(entry)) {
      // Don't add it to the backup
    } else if (entryStats.isDirectory()) {
      zip.addLocalFolder(entry, entry);
    } else {
      zip.addLocalFile(entry);
    }
  });

  fsExtra.ensureDirSync(path.dirname(outputPath));
  if (await zip.writeZipPromise(outputPath)) {
    console.log(`A backup of your project was created to ${outputPath}`);
  } else {
    throw new Error(`Error writting zip file to ${outputPath}`);
  }
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

async function updateManifestXmlReferences(): Promise<void> {
  const packageJson = `./package.json`;
  const readFileAsync = util.promisify(fs.readFile);
  const data = await readFileAsync(packageJson, "utf8");
  let content = JSON.parse(data);

  // Change .xml references to .json
  Object.keys(content.scripts).forEach(function (key) {
    content.scripts[key] = content.scripts[key].replace(
      /manifest.xml/gi,
      `manifest.json`
    );
  });
  // write updated json to file
  const writeFileAsync = util.promisify(fs.writeFile);
  await writeFileAsync(packageJson, JSON.stringify(content, null, 2));
}
