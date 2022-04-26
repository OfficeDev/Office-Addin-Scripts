// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import { ManifestUtil, TeamsAppManifest } from "@microsoft/teams-manifest";
import { v4 as uuidv4 } from "uuid";
import { ManifestInfo } from "../manifestInfo";
import { ManifestHandler } from "./manifestHandler";

export class ManifestHandlerJson extends ManifestHandler {
  /* eslint-disable @typescript-eslint/no-unused-vars */
  async modifyManifest(guid?: string, displayName?: string): Promise<TeamsAppManifest> {
    try {
      const teamsAppManifest: TeamsAppManifest = await ManifestUtil.loadFromPath(this.manifestPath);

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
    } catch (err) {
      throw new Error(`Unable to modify json data for manifest file: ${this.manifestPath}. \n${err}`);
    }
  }

  async parseManifest(): Promise<ManifestInfo> {
    try {
      const file: TeamsAppManifest = await ManifestUtil.loadFromPath(this.manifestPath);
      return this.getManifestInfo(file);
    } catch (err) {
      throw new Error(`Unable to read data for manifest file: ${this.manifestPath}. \n${err}`);
    }
  }

  // getManifestInfo(teamsAppManifest: TeamsAppManifest): ManifestInfo {
  getManifestInfo(teamsAppManifest: any): ManifestInfo {
    const manifestInfo: ManifestInfo = new ManifestInfo();
    manifestInfo.id = teamsAppManifest.id;
    // manifestInfo.appDomains = teamsAppManifest.validDomains; // Don't know what this is
    manifestInfo.defaultLocale = teamsAppManifest.localizationInfo?.defaultLanguageTag;
    manifestInfo.description = teamsAppManifest.description.short;
    manifestInfo.displayName = teamsAppManifest.name.short;
    manifestInfo.highResolutionIconUrl = teamsAppManifest.icons.color;
    manifestInfo.hosts = teamsAppManifest.extension?.runtimes[0]?.requirements?.capabilities?.map((capability: any) => capability.name); // We need for the new schema to be published
    manifestInfo.iconUrl = teamsAppManifest.icons.color;
    manifestInfo.officeAppType = teamsAppManifest.extension?.requirements?.capabilities[0]?.name; // We need for the new schema to be published
    manifestInfo.permissions = teamsAppManifest.authorization?.permissions?.resourceSpecific[0]?.name; // Still need to find an example where Permissions work. Or the manifest doesn't have it, or it is a wrong vlaue.
    manifestInfo.providerName = teamsAppManifest.developer.name;
    manifestInfo.supportUrl = teamsAppManifest.developer.websiteUrl;
    manifestInfo.version = teamsAppManifest.version;

    return manifestInfo;
  }

  async writeManifestData(manifestData: TeamsAppManifest): Promise<void> {
    await ManifestUtil.writeToPath(this.manifestPath, manifestData);
  }
  /* eslint-enable @typescript-eslint/no-unused-vars */
}
