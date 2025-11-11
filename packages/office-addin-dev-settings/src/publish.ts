// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import childProcess from "child_process";
import fs from "fs";

/* global console */

export type AccountOperation = "login" | "logout";

export async function registerWithTeams(filePath: string): Promise<string> {
  return new Promise((resolve, reject) => {
    if ((filePath.endsWith(".zip") || filePath.endsWith(".xml")) && fs.existsSync(filePath)) {
      const pathSwitch = filePath.endsWith(".zip") ? "--file-path" : "--xml-path";
      const sideloadCommand = `npx -p @microsoft/m365agentstoolkit-cli atk install ${pathSwitch} "${filePath}" --interactive false`;

      console.log(`running: ${sideloadCommand}`);
      childProcess.exec(sideloadCommand, (error, stdout, stderr) => {
        let titleIdMatch = stdout.match(/TitleId:\s*(.*)/);
        let titleId = titleIdMatch !== null ? titleIdMatch[1] : "??";
        if (error || stderr.match('"error"')) {
          console.log(`\n${stdout}\n--Error sideloading!--\nError: ${error}\nSTDERR:\n${stderr}`);
          reject(error);
        } else {
          console.log(
            `\n${stdout}\nSuccessfully registered package! (${titleId})\n STDERR: ${stderr}\n`
          );
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
    const authCommand = `npx -p @microsoft/m365agentstoolkit-cli atk auth ${operation} m365`;

    console.log(`running: ${authCommand}`);
    childProcess.exec(authCommand, (error, stdout, stderr) => {
      if (error || (stderr.length > 0 && /Debugger attached\./.test(stderr) == false)) {
        console.log(
          `Error running auth command\n STDOUT: ${stdout}\n ERROR: ${error}\n STDERR: ${stderr}`
        );
        reject(error);
      } else {
        console.log(`Successfully ran auth command.\n`);
        resolve();
      }
    });
  });
}

export async function uninstallWithTeams(id: string): Promise<boolean> {
  return new Promise((resolve, reject) => {
    const guidRegex = /[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}/;
    const manifestIdRegex = new RegExp(`^${guidRegex.source}$`);
    const titleIdRegex = new RegExp(`^U_${guidRegex.source}$`);
    let mode: string = "";

    if (titleIdRegex.test(id)) {
      mode = `--mode title-id --title-id ${id}`;
    } else if (manifestIdRegex.test(id)) {
      mode = `--mode manifest-id --manifest-id ${id}`;
    } else {
      console.error(`Error: Invalid id "${id}".  Add-in is still installed.`);
      resolve(false);
      return;
    }

    const uninstallCommand = `npx -p @microsoft/m365agentstoolkit-cli atk uninstall ${mode} --interactive false`;
    console.log(`running: ${uninstallCommand}`);
    childProcess.exec(uninstallCommand, (error, stdout, stderr) => {
      if (error || stderr.match('"error"')) {
        console.log(`\n${stdout}\n--Error uninstalling!--\n${error}\n STDERR: ${stderr}`);
        reject(error);
      } else {
        console.log(`\n${stdout}\nSuccessfully uninstalled!\n STDERR: ${stderr}\n`);
        resolve(true);
      }
    });
  });
}
