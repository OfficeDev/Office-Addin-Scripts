// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as childProcess from "child_process";
import * as fs from "fs";
import * as path from "path";

/* global console, process */

export type AccountOperation = "login" | "logout";

export async function registerWithTeams(zipPath: string): Promise<string> {
  return new Promise((resolve, reject) => {
    if (zipPath.endsWith(".zip") && fs.existsSync(zipPath)) {
      const sideloadCommand = `npx @microsoft/teamsfx-cli m365 sideloading --file-path ${zipPath}`;

      console.log(`running: ${sideloadCommand}`);
      childProcess.exec(sideloadCommand, (error, stdout, stderr) => {
        let titleIdMatch = stdout.match(/TitleId:\s*(.*)/);
        let titleId = titleIdMatch !== null ? titleIdMatch[1] : "??";
        if (error || stderr.match('"error"')) {
          console.log(`\n${stdout}\n--Error sideloading!--\nError: ${error}\nSTDERR:\n${stderr}`);
          reject(error);
        } else {
          console.log(`\n${stdout}\nSuccessfully registered package! (${titleId})\n STDERR: ${stderr}\n`);
          forceCacheUpdate();
          resolve(titleId);
        }
      });
    } else {
      reject(new Error(`The file '${zipPath}' is not valid`));
    }
  });
}

export async function updateM365Account(operation: AccountOperation): Promise<void> {
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

export async function unacquireWithTeams(titleId: string): Promise<void> {
  return new Promise((resolve, reject) => {
    const unacquireCommand = `npx @microsoft/teamsfx-cli m365 unacquire --title-id ${titleId}`;

    console.log(`running: ${unacquireCommand}`);
    childProcess.exec(unacquireCommand, (error, stdout, stderr) => {
      if (error || stderr.match('"error"')) {
        console.log(`\n${stdout}\n--Error unacquireing!--\n${error}\n STDERR: ${stderr}`);
        reject(error);
      } else {
        console.log(`\n${stdout}\nSuccessfully unacquired title!\n STDERR: ${stderr}\n`);
        forceCacheUpdate();
        resolve();
      }
    });
  });
}

function forceCacheUpdate() {
  // TODO: find HubAppFileCache on Mac and do the same targeted delete
  if (process.platform === "win32") {
    const cachePath: string = path.join(process.env.LOCALAPPDATA as string, "Microsoft\\Outlook\\HubAppFileCache");

    if (fs.existsSync(cachePath)) {
      // Get list of folders with hashed names
      const subFolders: fs.Dirent[] = fs.readdirSync(cachePath, { withFileTypes: true }).filter((entry: fs.Dirent) => {
        return entry.isDirectory();
      });

      // Delete any found file that prevents TAOS service calls
      subFolders.forEach((folder: fs.Dirent) => {
        const targetFiles: string[] = [
          path.resolve(cachePath, folder.name, "TaosSource\\CacheProperties"),
          path.resolve(cachePath, folder.name, "TaosSource\\PersistedCacheSynced"),
          path.resolve(cachePath, folder.name, "TaosSource\\u8qUM7HfoAFQ6YiuZO0RVQ=="),
          path.resolve(cachePath, folder.name, "TaosSource\\ZplQ1yfT07QnEV2xMoi+GQ=="),
        ];
        targetFiles.forEach((file) => {
          if (fs.existsSync(file)) {
            console.log(`Deleting File: ${file}`);
            fs.unlinkSync(file);
          }
        });
      });
    }
  }
}
