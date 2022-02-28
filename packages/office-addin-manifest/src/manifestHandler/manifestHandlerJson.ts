// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import { TeamsAppManifest } from "@microsoft/teamsfx-api";
import { ManifestUtil } from "@microsoft/teams-manifest";
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
      return this.getManifestInfo(file);
    } catch (err) {
      throw new Error(`Unable to read data for manifest file: ${this.manifestPath}. \n${err}`);
    }
  }

  getManifestInfo(teamsAppManifest: TeamsAppManifest): ManifestInfo {
    const manifestInfo: ManifestInfo = new ManifestInfo();
    manifestInfo.id = teamsAppManifest.id;
    manifestInfo.appDomains = teamsAppManifest.validDomains;
    manifestInfo.defaultLocale = teamsAppManifest.localizationInfo?.defaultLanguageTag;
    manifestInfo.description = teamsAppManifest.description.short;
    manifestInfo.displayName = teamsAppManifest.name.short;
    manifestInfo.highResolutionIconUrl = teamsAppManifest.icons.color;
    // manifestInfo.hosts = teamsAppManifest.extension.requirements.scopes; // We need for the new schema to be published
    manifestInfo.iconUrl = teamsAppManifest.icons.color;
    // manifestInfo.officeAppType = teamsAppManifest.extension.requirements.capabilities.name; // We need for the new schema to be published
    // manifestInfo.permissions = teamsAppManifest.Permissions; // Still need to find an example where Permissions work. Or the manifest doesn't have it, or it is a wrong vlaue.
    manifestInfo.providerName = teamsAppManifest.developer.name;
    manifestInfo.supportUrl = teamsAppManifest.developer.websiteUrl;
    manifestInfo.version = teamsAppManifest.version;

    return manifestInfo;
  }

  async writeManifestData(manifestData: TeamsAppManifest): Promise<void> {
    await ManifestUtil.writeToPath(this.manifestPath, manifestData);
    await this.readFromManifestFile();
  }
  /* eslint-enable @typescript-eslint/no-unused-vars */
}
