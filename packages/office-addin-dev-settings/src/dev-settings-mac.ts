// copyright (c) Microsoft Corporation. All rights reserved.
// licensed under the MIT license.

import * as fs from "fs-extra";
import * as junk from "junk";
import {
  exportMetadataPackage,
  getOfficeApps,
  getOfficeAppsForManifestHosts,
  OfficeApp,
  OfficeAddinManifest,
} from "office-addin-manifest";
import * as os from "os";
import * as path from "path";
import { RegisteredAddin } from "./dev-settings";
import { ExpectedError } from "office-addin-usage-data";
import { registerWithTeams, unacquireWithTeams } from "./publish";
import * as fspath from "path";

/* global process */

export async function getRegisteredAddIns(): Promise<RegisteredAddin[]> {
  const registeredAddins: RegisteredAddin[] = [];

  for (const app of getOfficeApps()) {
    const sideloadDirectory = getSideloadDirectory(app);

    if (sideloadDirectory && fs.existsSync(sideloadDirectory)) {
      for (const fileName of fs.readdirSync(sideloadDirectory).filter(junk.not)) {
        const manifestPath = fs.realpathSync(path.join(sideloadDirectory, fileName));
        const manifest = await OfficeAddinManifest.readManifestFile(manifestPath);
        registeredAddins.push(new RegisteredAddin(manifest.id || "", manifestPath));
      }
    }
  }

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

export async function registerAddIn(manifestPath: string, officeApps?: OfficeApp[]) {
  try {
    const manifest = await OfficeAddinManifest.readManifestFile(manifestPath);
    let key = path.basename(manifestPath);

    if (!officeApps) {
      officeApps = getOfficeAppsForManifestHosts(manifest.hosts);

      if (officeApps.length === 0) {
        throw new ExpectedError("The manifest file doesn't specify any hosts for the Office Add-in.");
      }
    }

    if (!manifest.id) {
      throw new ExpectedError("The manifest file doesn't contain the id of the Office Add-in.");
    }

    if (manifestPath.endsWith(".json")) {
      const targetPath: string = fspath.join(process.env.TEMP as string, "manifest.zip");
      const zipPath: string = await exportMetadataPackage(targetPath, manifestPath);
      key = await registerWithTeams(zipPath);
    }
    for (const app of officeApps) {
      const sideloadDirectory = getSideloadDirectory(app);

      if (sideloadDirectory) {
        // include manifest id in sideload filename
        const sideloadPath = path.join(sideloadDirectory, `${manifest.id}.${key}`);

        fs.ensureDirSync(sideloadDirectory);
        fs.ensureLinkSync(manifestPath, sideloadPath);
      }
    }
  } catch (err) {
    throw new Error(`Unable to register the Office Add-in.\n${err}`);
  }
}

export async function unregisterAddIn(addinId: string, manifestPath: string): Promise<void> {
  if (!addinId) {
    throw new ExpectedError("The manifest file doesn't contain the id of the Office Add-in.");
  }

  const registeredAddIns = await getRegisteredAddIns();

  for (const registeredAddIn of registeredAddIns) {
    const registeredFileName = path.basename(registeredAddIn.manifestPath);
    const manifestFileName = path.basename(manifestPath);
    
    if (registeredFileName === manifestFileName || registeredFileName.startsWith(addinId)) {
      if(!registeredFileName.endsWith("*.xml")) {
        unacquireWithTeams(registeredFileName.substring(registeredFileName.indexOf(".")));
      }
      fs.unlinkSync(registeredAddIn.manifestPath);
    }
  }
}

export async function unregisterAllAddIns(): Promise<void> {
  const registeredAddIns = await getRegisteredAddIns();

  for (const registeredAddIn of registeredAddIns) {
    const registeredFileName = path.basename(registeredAddIn.manifestPath);
    if(!registeredFileName.endsWith("*.xml")) {
      unacquireWithTeams(registeredFileName.substring(registeredFileName.indexOf(".")));
    }
    fs.unlinkSync(registeredAddIn.manifestPath);
  }
}
