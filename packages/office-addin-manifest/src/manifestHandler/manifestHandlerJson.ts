// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import { devPreview, ManifestUtil } from "@microsoft/teams-manifest";
import { v4 as uuidv4 } from "uuid";
import { ManifestInfo, ManifestType } from "../manifestInfo";
import { ManifestHandler } from "./manifestHandler";

export class ManifestHandlerJson extends ManifestHandler {
  /* eslint-disable @typescript-eslint/no-unused-vars */
  async modifyManifest(guid?: string, displayName?: string): Promise<devPreview.DevPreviewSchema> {
    try {
      const appManifest: devPreview.DevPreviewSchema = await ManifestUtil.loadFromPath(this.manifestPath);

      if (typeof guid !== "undefined") {
        if (!guid || guid === "random") {
          guid = uuidv4();
        }
        appManifest.id = guid;
      }

      if (typeof displayName !== "undefined") {
        appManifest.name.short = displayName;
      }
      return appManifest;
    } catch (err) {
      throw new Error(`Unable to modify json data for manifest file: ${this.manifestPath}. \n${err}`);
    }
  }

  async parseManifest(): Promise<ManifestInfo> {
    try {
      const file: devPreview.DevPreviewSchema = await ManifestUtil.loadFromPath(this.manifestPath);
      return this.getManifestInfo(file);
    } catch (err) {
      throw new Error(`Unable to read data for manifest file: ${this.manifestPath}. \n${err}`);
    }
  }

  getManifestInfo(appManifest: devPreview.DevPreviewSchema): ManifestInfo {
    const manifestInfo: ManifestInfo = new ManifestInfo();

    // Need to handle mutliple version of the prerelease json manifest schema
    const extensionElement = appManifest.extensions?.[0];

    manifestInfo.id = appManifest.id;
    manifestInfo.appDomains = appManifest.validDomains;
    manifestInfo.defaultLocale = appManifest.localizationInfo?.defaultLanguageTag;
    manifestInfo.description = appManifest.description.short;
    manifestInfo.displayName = appManifest.name.short;
    manifestInfo.highResolutionIconUrl = appManifest.icons.color;
    manifestInfo.hosts = extensionElement?.requirements?.scopes;
    manifestInfo.iconUrl = appManifest.icons.color;
    manifestInfo.officeAppType = extensionElement?.requirements?.capabilities?.[0]?.name;
    manifestInfo.permissions = appManifest.authorization?.permissions?.resourceSpecific?.[0]?.name;
    manifestInfo.providerName = appManifest.developer.name;
    manifestInfo.supportUrl = appManifest.developer.websiteUrl;
    manifestInfo.version = appManifest.version;
    manifestInfo.manifestType = ManifestType.JSON;

    return manifestInfo;
  }

  async writeManifestData(manifestData: devPreview.DevPreviewSchema): Promise<void> {
    await ManifestUtil.writeToPath(this.manifestPath, manifestData);
  }
  /* eslint-enable @typescript-eslint/no-unused-vars */
}
