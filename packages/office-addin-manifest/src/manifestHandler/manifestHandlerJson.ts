// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import { ManifestInfo } from "../manifestInfo";
import { ManifestHandler } from "./manifestHandler";

export class ManifestHandlerJson extends ManifestHandler {
  /* eslint-disable no-unused-vars */
  async modifyManifest(manifestPath: string, fileData: string, guid?: string, displayName?: string): Promise<any> {
    console.log("modifyManifest JSON called");
  }

  async parseManifest(manifestPath: string, fileData: string): Promise<ManifestInfo> {
    console.log("parseManifest JSON called");
    const manifestInfo: ManifestInfo = new ManifestInfo();

    try {
      const file = JSON.parse(fileData);
    } catch (err) {
      throw new Error(`Unable to read data for manifest file2: ${manifestPath}. \n${err}`);
    }

    return manifestInfo;
  }

  async writeManifestData(manifestPath: string, manifestData: any): Promise<void> {
    console.log("writeManifestData JSON called");

  }
  /* eslint-enable no-unused-vars */
}
