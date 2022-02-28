// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import { ManifestInfo } from "../manifestInfo";

export abstract class ManifestHandler {
  /* eslint-disable no-unused-vars */
  constructor(manifestPath: string) {
    this.manifestPath = manifestPath;
    this.fileData = "";
  }

  abstract modifyManifest(guid?: string, displayName?: string): Promise<any>;
  abstract parseManifest(): Promise<ManifestInfo>;
  abstract writeManifestData(manifestData: any): Promise<void>;
  manifestPath: string;
  fileData: string;
  /* eslint-enable no-unused-vars */
}
