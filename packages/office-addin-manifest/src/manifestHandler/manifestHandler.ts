// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import { ManifestInfo } from "../manifestInfo";

export abstract class ManifestHandler {
  constructor(manifestPath: string) {
    this.manifestPath = manifestPath;
  }

  abstract modifyManifest(guid?: string, displayName?: string): Promise<any>;
  abstract parseManifest(): Promise<ManifestInfo>;
  abstract writeManifestData(manifestData: any): Promise<void>;
  manifestPath: string;
}
