// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as fs from "fs-extra";
import { TeamsAppManifest } from "@microsoft/teamsfx-api";
import { v4 as uuidv4 } from "uuid";
import { ManifestInfo } from "../manifestInfo";
import { ManifestHandler } from "./manifestHandler";

export class ManifestHandlerJson extends ManifestHandler {
  /* eslint-disable @typescript-eslint/no-unused-vars */
  async modifyManifest(guid?: string, displayName?: string): Promise<TeamsAppManifest> {
    const teamsAppManifest: TeamsAppManifest = await JSON.parse(this.fileData);

    if (typeof guid !== "undefined") {
      if (!guid || guid === "random") {
        guid = uuidv4();
      }
      teamsAppManifest.id = guid;
    }

    if (typeof displayName !== "undefined") {
      teamsAppManifest.name.short = displayName;
    }
    return teamsAppManifest;
  }

  async parseManifest(): Promise<ManifestInfo> {
    try {
      const file: TeamsAppManifest = await JSON.parse(this.fileData);
      return this.translateTeamsAppManifestToManifestInfo(file);
    } catch (err) {
      throw new Error(`Unable to read data for manifest file: ${this.manifestPath}. \n${err}`);
    }
  }

  translateTeamsAppManifestToManifestInfo(teamsAppManifest: TeamsAppManifest): ManifestInfo {
    const manifestInfo: ManifestInfo = new ManifestInfo();
    manifestInfo.id = teamsAppManifest.id;
    // manifestInfo.allowSnapshot = teamsAppManifest.AllowSnapshot; // TODO
    // manifestInfo.alternateId = teamsAppManifest.AlternateId; // TODO
    manifestInfo.appDomains = teamsAppManifest.validDomains;
    manifestInfo.defaultLocale = teamsAppManifest.localizationInfo?.defaultLanguageTag;
    manifestInfo.description = teamsAppManifest.description.short;
    manifestInfo.displayName = teamsAppManifest.name.short;
    manifestInfo.highResolutionIconUrl = teamsAppManifest.icons.color;
    // manifestInfo.hosts = teamsAppManifest.Hosts.Host.Name; // TODO
    manifestInfo.iconUrl = teamsAppManifest.icons.color; // TODO
    // manifestInfo.officeAppType = teamsAppManifest["xsi:type"]; // TODO
    // manifestInfo.permissions = teamsAppManifest.Permissions; // TODO
    manifestInfo.providerName = teamsAppManifest.developer.name;
    manifestInfo.supportUrl = teamsAppManifest.developer.websiteUrl; // TODO
    manifestInfo.version = teamsAppManifest.version;

    return manifestInfo;
  }

  async writeManifestData(manifestData: TeamsAppManifest): Promise<void> {
    await writeToPath(this.manifestPath, manifestData);
    await this.updateFileData();
  }
  /* eslint-enable @typescript-eslint/no-unused-vars */
}

/**
 * Save manifest to .json file
 * @param filePath path to the manifest.json file
 */
async function writeToPath(path: string, manifest: TeamsAppManifest): Promise<void> {
  return fs.writeJson(path, manifest, { spaces: 4 });
}
