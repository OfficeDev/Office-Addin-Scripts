// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as childProcess from "child_process";
import * as fs from "fs";
import * as path from "path";

/* global console, __dirname */

export async function registerWithTeams(zipPath: string): Promise<void> {
  return new Promise((resolve, reject) => {
    if (zipPath.endsWith(".zip") && fs.existsSync(zipPath)) {
      const cliPath = path.resolve(__dirname, "..\\node_modules\\@microsoft\\teamsfx-cli\\cli.js");
      const configPath = path.resolve(__dirname, ".\\sdf.json");
      const loginCommand = `node ${cliPath} account login m365`;
      const sideloadCommand = `node ${cliPath} m365 sideloading --file-path ${zipPath} --service-config ${configPath}`;

      console.log(`running: ${loginCommand}`);
      childProcess.exec(loginCommand, (error, stdout, stderr) => {
        if (error || stderr.length > 0) {
          console.log(`Error logging in:\n STDOUT: ${stdout}\n ERROR: ${error}\n STDERR: ${stderr}`);
          reject(error);
        } else {
          console.log(`Successfully logged in.\n`);
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
          resolve();
        }
      });
    } else {
      reject(new Error(`The file '${zipPath}' is not valid`));
    }
  });
}
