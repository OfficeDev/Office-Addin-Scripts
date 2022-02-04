// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as fs from "fs";
import * as semver from "semver";
import { ExpectedError } from "office-addin-usage-data";

export async function convertProject(
  manifestPath: string = "./manifest.xml",
  packageJsonPath: string = "./package.json"
) {
  if (manifestPath.endsWith(".json")) {
    throw new ExpectedError(`The convert command only works on xml based projects`);
  }

  if (!fs.existsSync(manifestPath)) {
    throw new ExpectedError(`The manifest file '${manifestPath}' does not exist`);
  }

  if (!fs.existsSync(packageJsonPath)) {
    throw new ExpectedError(`The package.json file '${packageJsonPath}' does not exist`);
  }

  checkPackagesAreUpdated(packageJsonPath);
}

function checkPackagesAreUpdated(packageJsonPath: string): void {
  const data = fs.readFileSync(packageJsonPath, "utf8");
  let content = JSON.parse(data);

  // Contains name of the package and minimum version
  const depedentPackages: Map<string, string> = new Map<string, string>([
    ["office-addin-debugging", "4.3.7"],
    ["office-addin-manifest", "1.7.6"],
  ]);

  Object.keys(content.devDependencies).forEach(function (key) {
    if (depedentPackages.has(key)) {
      const minVersion: string = depedentPackages.get(key) ?? "";
      const version = semver.coerce(content.devDependencies[key]);

      if (version && !semver.gte(version, minVersion)) {
        throw new ExpectedError(
          `Your version of the package ${key} should be at least ${depedentPackages.get(
            key
          )}. To update your package, run:\nnpm i ${key}@latest --save-dev`
        );
      }
    }
  });
}
