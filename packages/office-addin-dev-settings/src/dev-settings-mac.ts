// copyright (c) Microsoft Corporation. All rights reserved.
// licensed under the MIT license.

import { RegisteredAddin } from "./dev-settings";
import * as junk from "junk";
import { OfficeApp, getOfficeApps } from "./officeApp";
import * as os from "os";
import * as path from "path";
import * as fs from "fs";

export async function getRegisteredAddIns(): Promise<RegisteredAddin[]> {
  let registeredAddins: RegisteredAddin[] = [];

  getOfficeApps().forEach(app => {
    const sideloadDirectory = getSideloadDirectory(app);

    if (sideloadDirectory && fs.existsSync(sideloadDirectory)) {
      fs.readdirSync(sideloadDirectory).filter(junk.not).forEach(fileName => {
        const manifestPath = fs.realpathSync(path.join(sideloadDirectory, fileName));        
        registeredAddins.push(new RegisteredAddin("", manifestPath));
      });
    }
  });

  return registeredAddins;
}

function getSideloadDirectory(app: OfficeApp): string | undefined {
  switch (app) {
    case OfficeApp.Excel:
      return path.join(os.homedir(), "Library/Containers/com.microsoft.Excel/Data/Documents/wef");
    case OfficeApp.PowerPoint:
      return path.join(os.homedir(), "Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef");
    case OfficeApp.Word:
      return path.join(os.homedir(), "Library/Containers/com.microsoft.Word/Data/Documents/wef");
  }
}

export async function registerAddIn(manifestPath: string) {
}

export async function unregisterAddIn(manifestPath: string): Promise<void> {
}

export async function unregisterAllAddIns(): Promise<void> {
}