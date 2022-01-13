// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import { ManifestInfo } from "../manifestInfo";
import { ManifestHandler } from "./manifestHandler";

export class ManifestHandlerJson extends ManifestHandler {
  /* eslint-disable no-unused-vars */
  async modifyManifest(manifestPath: string, fileData: string, guid?: string, displayName?: string): Promise<any> {}

  async parseManifest(manifestPath: string, fileData: string): Promise<ManifestInfo> {
    const manifest: ManifestInfo = new ManifestInfo();
    return manifest;
  }

  async writeManifestData(manifestPath: string, manifestData: any): Promise<void> {}
  /* eslint-enable no-unused-vars */
}
