// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import fs from "fs";
import fspath from "path";

export function getDiskManifestDir(create: boolean): string {
  const tempDir = process.env.TEMP;
  const targetDirName = "OfficeAddinDebugging";
  const targetManifestDir = fspath.normalize(`${tempDir}/${targetDirName}`);

  if (create && !fs.existsSync(targetManifestDir)) {
    fs.mkdirSync(targetManifestDir);
  }
  return targetManifestDir;
}