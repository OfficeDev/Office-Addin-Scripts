// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import { ManifestInfo } from "../manifestInfo";

export abstract class ManifestHandler {
  /* eslint-disable no-unused-vars */
  abstract modifyManifest(manifestPath: string, guid?: string, displayName?: string): Promise<any>;
  abstract parseManifest(file: any): ManifestInfo;
  abstract readFromManifestFile(manifestPath: string): Promise<any>;
  abstract writeManifestData(manifestPath: string, manifestData: any): Promise<void>;
  /* eslint-enable no-unused-vars */
}
