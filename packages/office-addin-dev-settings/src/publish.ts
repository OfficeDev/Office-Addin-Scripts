// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as childProcess from "child_process";
import * as fs from "fs";
import * as path from "path";

/* global console */

export async function registerWithTeams(zipPath: string): Promise<void> {
  return new Promise((resolve, reject) => {
    if (zipPath.endsWith(".zip") && fs.existsSync(zipPath)) {
      const sideloadCommand = `npx @microsoft/teamsfx-cli m365 sideloading --file-path ${zipPath}`;

      console.log(`running: ${sideloadCommand}`);
      childProcess.exec(sideloadCommand, (error, stdout, stderr) => {
        if (error || stderr.length > 0) {
          console.log(`Error sideloading:\n STDOUT: ${stdout}\n ERROR: ${error}\n STDERR: ${stderr}`);
          reject(error);
        } else {
          console.log(`Successfully registered package:\n ${stdout}\n`);
          resolve();
        }
      });
    } else {
      reject(new Error(`The file '${zipPath}' is not valid`));
    }
  });
}

export async function updateM365Account(operation: "login" | "logout"): Promise<void> {
  return new Promise((resolve, reject) => {
    const loginCommand = `npx @microsoft/teamsfx-cli account ${operation} m365`;

    console.log(`running: ${loginCommand}`);
    childProcess.exec(loginCommand, (error, stdout, stderr) => {
      if (error || (stderr.length > 0 && /Debugger attached\./.test(stderr) == false)) {
        console.log(`Error logging in:\n STDOUT: ${stdout}\n ERROR: ${error}\n STDERR: ${stderr}`);
        reject(error);
      } else {
        console.log(`Successfully logged in/out.\n`);
        resolve();
      }
    });
  });
}
