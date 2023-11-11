// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as childProcess from "child_process";
import * as fs from "fs";
import * as path from "path";

/* global console, process */

export type AccountOperation = "login" | "logout";

export async function registerWithTeams(filePath: string): Promise<string> {
  return new Promise((resolve, reject) => {
    if ((filePath.endsWith(".zip") || filePath.endsWith(".xml")) && fs.existsSync(filePath)) {
      const pathSwitch = filePath.endsWith(".zip") ? "--file-path" : "--xml-path";
      const sideloadCommand = `npx -p @microsoft/teamsfx-cli@2.0.3-alpha.e5db52f29.0 teamsfx m365 sideloading ${pathSwitch} "${filePath}"`;

      console.log(`running: ${sideloadCommand}`);
      childProcess.exec(sideloadCommand, (error, stdout, stderr) => {
        let titleIdMatch = stdout.match(/TitleId:\s*(.*)/);
        let titleId = titleIdMatch !== null ? titleIdMatch[1] : "??";
        if (error || stderr.match('"error"')) {
          console.log(`\n${stdout}\n--Error sideloading!--\nError: ${error}\nSTDERR:\n${stderr}`);
          reject(error);
        } else {
          console.log(`\n${stdout}\nSuccessfully registered package! (${titleId})\n STDERR: ${stderr}\n`);
          resolve(titleId);
        }
      });
    } else {
      reject(new Error(`The file '${filePath}' is not valid`));
    }
  });
}

export async function updateM365Account(operation: AccountOperation): Promise<void> {
  return new Promise((resolve, reject) => {
    const loginCommand = `npx -p @microsoft/teamsfx-cli@2.0.3-alpha.e5db52f29.0 teamsfx account ${operation} m365`;

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

export async function unacquireWithTeams(titleId: string): Promise<void> {
  return new Promise((resolve, reject) => {
    const unacquireCommand = `npx -p @microsoft/teamsfx-cli@2.0.3-alpha.e5db52f29.0 teamsfx m365 unacquire --title-id ${titleId}`;

    console.log(`running: ${unacquireCommand}`);
    childProcess.exec(unacquireCommand, (error, stdout, stderr) => {
      if (error || stderr.match('"error"')) {
        console.log(`\n${stdout}\n--Error unacquireing!--\n${error}\n STDERR: ${stderr}`);
        reject(error);
      } else {
        console.log(`\n${stdout}\nSuccessfully unacquired title!\n STDERR: ${stderr}\n`);
        resolve();
      }
    });
  });
}
