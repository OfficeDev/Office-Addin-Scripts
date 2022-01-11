// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import { ManifestInfo } from "../manifestInfo";
import { ManifestHandler } from "./manifestHandler";

export class ManifestHandlerJson extends ManifestHandler {
  async modifyManifest(manifestPath: string, guid?: string, displayName?: string): Promise<any> {}

  parseManifest(file: any): ManifestInfo {
    const manifest: ManifestInfo = new ManifestInfo();
    return manifest;
  }

  async readFromManifestFile(manifestPath: string): Promise<any> {
    return "json data";
  }

  async writeManifestData(manifestPath: string, manifestData: any): Promise<void> {}
}
