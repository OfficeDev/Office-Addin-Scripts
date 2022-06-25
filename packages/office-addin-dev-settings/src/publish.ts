// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as childProcess from "child_process";
import * as fs from "fs";

/* global console */

export async function registerWithTeams(zipPath: string): Promise<void> {
  return new Promise((resolve, reject) => {
    if (zipPath.endsWith(".zip") && fs.existsSync(zipPath)) {
      const command = `npx @microsoft/teamsfx-cli provision manifest --file-path ${zipPath}`;

      childProcess.exec(command, (error, stdout) => {
        if (error) {
          reject(error);
        } else {
          console.log("Successfully registered package with https://dev.teams.microsoft.com/apps");
          resolve();
        }
      });
    } else {
      reject(new Error(`The file '${zipPath}' is not valid`));
    }
  });
}
