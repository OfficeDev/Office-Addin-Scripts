// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import { ManifestInfo } from "../manifestInfo";
import { ManifestHandler } from "./manifestHandler";

export class ManifestHandlerJson extends ManifestHandler {
  /* eslint-disable no-unused-vars */
  async modifyManifest(manifestPath: string, fileData: string, guid?: string, displayName?: string): Promise<any> {
    throw new Error("Manifest cannot be modified in .json files");
  }

  async parseManifest(manifestPath: string, fileData: string): Promise<ManifestInfo> {
    const manifestInfo: ManifestInfo = new ManifestInfo();

    try {
      const file = await JSON.parse(fileData);
    } catch (err) {
      throw new Error(`Unable to read data for manifest file: ${manifestPath}. \n${err}`);
    }

    throw new Error("Manifest cannot be parsed in .json files");
  }

  async writeManifestData(manifestPath: string, manifestData: any): Promise<void> {
    throw new Error("Manifest cannot be written in .json files");
  }
  /* eslint-enable no-unused-vars */
}
