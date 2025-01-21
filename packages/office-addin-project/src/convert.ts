// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import AdmZip from "adm-zip";
import fs from "fs";
import fsExtra from "fs-extra";
import path from "path";
import util from "util";
import { exec } from "child_process";
import { convert } from "office-addin-manifest-converter";
import { ExpectedError } from "office-addin-usage-data";

const execAsync = util.promisify(exec);
const skipBackup: string[] = ["node_modules"];

export async function convertProject(
  manifestPath: string = "./manifest.xml",
  backupPath: string = "./backup.zip",
  projectDir: string = "",
  devPreview: boolean = false
) {
  if (manifestPath.endsWith(".json")) {
    throw new ExpectedError(`The convert command only works on xml manifest based projects`);
  }

  if (!fs.existsSync(manifestPath)) {
    throw new ExpectedError(`The manifest file '${manifestPath}' does not exist`);
  }

  const outputPath: string = path.dirname(manifestPath);
  const currentDir: string = process.cwd();

  await backupProject(backupPath);
  try {
    // assume project dir is the same as manifest dir if not specified
    projectDir = projectDir === "" ? outputPath : projectDir;
    process.chdir(projectDir);

    if (!devPreview) {
      await convert(manifestPath, outputPath, false /* imageDownload */);
    } else {
      // override the schema used in the json manifest to use the devPreview schema
      await convert(
        manifestPath,
        outputPath,
        false /* imageDownload */,
        "https://developer.microsoft.com/json-schemas/teams/vDevPreview/MicrosoftTeams.schema.json",
        "devPreview"
      );
    }
    await updatePackages();
    await updateManifestXmlReferences();
    fs.unlinkSync(manifestPath);
  } catch (err: any) {
    console.log(`Error in conversion. Restoring project initial state.`);
    await restoreBackup(backupPath);
    throw err;
  } finally {
    process.chdir(currentDir);
  }
}

async function backupProject(backupPath: string) {
  const zip: AdmZip = new AdmZip();
  const outputPath: string = path.resolve(backupPath);
  const rootDir: string = path.resolve();

  const files: string[] = fs.readdirSync(rootDir);
  files.forEach((entry) => {
    const fullPath = path.resolve(entry);
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

async function restoreBackup(backupPath: string) {
  var zip = new AdmZip(backupPath); // reading archives
  zip.extractAllTo("./", true); // overwrite
}

async function updatePackages(): Promise<void> {
  // Contains name of the package and minimum version
  const depedentPackages: string[] = ["office-addin-debugging", "office-addin-manifest"];
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
  await execAsync(command);
}

async function updateManifestXmlReferences(): Promise<void> {
  await updatePackageJson();
  await updateWebpackConfig();
}

async function updatePackageJson(): Promise<void> {
  const packageJson = `./package.json`;
  const readFileAsync = util.promisify(fs.readFile);
  const data = await readFileAsync(packageJson, "utf8");
  let content = JSON.parse(data);

  // Change .xml references to .json
  Object.keys(content.scripts).forEach(function (key) {
    content.scripts[key] = content.scripts[key].replace(/manifest.xml/gi, `manifest.json`);
  });

  // write updated json to file
  const writeFileAsync = util.promisify(fs.writeFile);
  await writeFileAsync(packageJson, JSON.stringify(content, null, 2));
}

async function updateWebpackConfig(): Promise<void> {
  const weppackConfig = `./webpack.config.js`;
  const readFileAsync = util.promisify(fs.readFile);
  const data = await readFileAsync(weppackConfig, "utf8");

  // switching to json extension is the easy fix.
  // TODO: update to remove the manifest copy plugin since it's not needed in webpack
  let content = data.replace(/"(manifest\*\.)xml"/gi, '"$1json"');

  const writeFileAsync = util.promisify(fs.writeFile);
  await writeFileAsync(weppackConfig, content);
}
