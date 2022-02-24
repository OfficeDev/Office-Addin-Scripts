// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import { ManifestInfo } from "../manifestInfo";
import * as util from "util";
import * as fs from "fs";

export abstract class ManifestHandler {
  /* eslint-disable no-unused-vars */
  constructor(manifestPath: string) {
    this.manifestPath = manifestPath;
    this.fileData = "";
  }

  public async readFromManifestFile(manifestPath: string): Promise<string> {
    try {
      const fileData: string = await util.promisify(fs.readFile)(manifestPath, {
        encoding: "utf8",
      });
      this.fileData = fileData;
      return fileData;
    } catch (err) {
      throw new Error(`Unable to read data for manifest file: ${manifestPath}. \n${err}`);
    }
  }

  abstract modifyManifest(guid?: string, displayName?: string): Promise<any>;
  abstract parseManifest(): Promise<ManifestInfo>;
  abstract writeManifestData(manifestData: any): Promise<void>;
  manifestPath: string;
  fileData: string;
  /* eslint-enable no-unused-vars */
}
