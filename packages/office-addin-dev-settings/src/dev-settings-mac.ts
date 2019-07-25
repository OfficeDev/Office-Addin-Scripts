// copyright (c) Microsoft Corporation. All rights reserved.
// licensed under the MIT license.

import { RegisteredAddin } from "./dev-settings";
import * as junk from "junk";
import { OfficeApp, getOfficeApps } from "./officeApp";
import * as os from "os";
import * as path from "path";
import * as fs from "fs-extra";
import * as officeAddinManifest from "office-addin-manifest";

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

function getOfficeAppForManifestHost(host: string): OfficeApp | undefined {
  switch (host.toLowerCase()) {
    case "document":
      return OfficeApp.Word;
    case "mailbox":
      return OfficeApp.Outlook;
    case "notebook":
      return OfficeApp.OneNote;
    case "presentation":
      return OfficeApp.PowerPoint;
    case "project":
      return OfficeApp.Project;
    case "workbook":
      return OfficeApp.Excel;
    default:
      return undefined;
  }
}  

function getOfficeAppsForManifestHosts(hosts?: string[]): OfficeApp[] {
  const officeApps: OfficeApp[] = [];

  if (hosts) {
    hosts.forEach(host => {
      const officeApp = getOfficeAppForManifestHost(host);

      if (officeApp) {
        officeApps.push(officeApp);
      }
    });
  }

  return officeApps;
}

export async function registerAddIn(manifestPath: string, officeApps?: OfficeApp[]) {
  try {
    const manifest = await officeAddinManifest.readManifestFile(manifestPath);

    if (!officeApps) {
      officeApps = getOfficeAppsForManifestHosts(manifest.hosts);

      if (officeApps.length === 0) {
        throw new Error("The manifest file doesn't specify any hosts for the Office add-in.")
      }
    }
 
    if (!manifest.id) {
      throw new Error("The manifest file doesn't contain the id of the Office add-in.");
    }

    officeApps.forEach(app => {
      const sideloadDirectory = getSideloadDirectory(app);

      if (sideloadDirectory) {
        const sideloadPath = path.join(sideloadDirectory, `${manifest.id}.${path.basename(manifestPath)}`);

        fs.ensureDirSync(sideloadDirectory);
        fs.ensureLinkSync(manifestPath, sideloadPath);
      }
    });
  } catch (err) {
    throw new Error(`Unable to register the Office add-in.\n${err}`);
  }
}

export async function unregisterAddIn(manifestPath: string): Promise<void> {
}

export async function unregisterAllAddIns(): Promise<void> {
}