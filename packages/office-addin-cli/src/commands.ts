// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as commander from "commander";
import * as fs from "fs";
import * as semver from "semver";
// import semver = require("semver");

import * as util from "util";
import { ExpectedError, logErrorMessage } from "office-addin-usage-data";
import { usageDataObject } from "./defaults";

export async function convert(command: commander.Command) {
  try {
    // Check if the project is already converted

    // Convert manifest .xml to manifest.json
    // Do any more needed modifications on the code?

    const manifestPath: string = command.manifest ?? "./manifest.xml";

    if (manifestPath.endsWith(".json")) {
      throw new ExpectedError(`The convert command only works on xml based manifests`);
    }

    if (!fs.existsSync(manifestPath)) {
      throw new ExpectedError(`The file '${manifestPath}' does not exist`);
    }
  
    const manifestData = fs.readFileSync(manifestPath);

    const validPackages: boolean = await checkPackagesAreUpdated();
    if (!validPackages) {

    }


    usageDataObject.reportSuccess("convert");
    throw new Error("Upgrade function is not ready yet.");
  } catch (err: any) {
    usageDataObject.reportException("convert", err);
    logErrorMessage(err);
  }
}

const readFileAsync = util.promisify(fs.readFile);
export async function checkPackagesAreUpdated(packageJsonPath?: string): Promise<boolean> {
  const packageJson = packageJsonPath ?? `./package.json`;
  const data = await readFileAsync(packageJson, "utf8");
  let content = JSON.parse(data);

  // Contains name of the package and minimum version
  const depedentPackages: Map<string, string> = new Map<string, string>([
    ["office-addin-debugging", "4.3.7"],
    ["office-addin-manifest", "1.7.6"],
  ]);

  let versionsAreValid = true;
  Object.keys(content.devDependencies).forEach(function (key) {
    if (depedentPackages.has(key)) {
      const minVersion: string = depedentPackages.get(key) ?? "";
      const version = semver.coerce(content.devDependencies[key]);

      if (version && !semver.gte(version, minVersion)) {
        versionsAreValid = false;
      }
    }
  });

  return versionsAreValid;
}
