// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import { ManifestInfo } from "../manifestInfo";
import { ManifestHandler } from "./manifestHandler";
import { TeamsAppManifest, IStaticTab, IConfigurableTab } from "@microsoft/teamsfx-api";
import * as fs from "fs-extra";

export class ManifestHandlerJson extends ManifestHandler {
  /* eslint-disable @typescript-eslint/no-unused-vars */
  async modifyManifest(guid?: string, displayName?: string): Promise<any> {
    throw new Error("Manifest cannot be modified in .json files");
  }

  async parseManifest(): Promise<ManifestInfo> {
    const manifestInfo: ManifestInfo = new ManifestInfo();

    try {
      const file = await JSON.parse(this.fileData);
    } catch (err) {
      throw new Error(`Unable to read data for manifest file: ${this.manifestPath}. \n${err}`);
    }

    throw new Error("Manifest cannot be parsed in .json files");
  }

  async writeManifestData(manifestData: any): Promise<void> {
    throw new Error("Manifest cannot be written in .json files");
  }
  /* eslint-enable @typescript-eslint/no-unused-vars */
}

/**
 * Save manifest to .json file
 * @param filePath path to the manifest.json file
 */
async function save(filePath: string, manifestData: any): Promise<void> {
  await fs.writeFile(filePath, JSON.stringify(manifestData, null, 4));
}

/**
* Load manifest from manifest.json
* @param filePath path to the manifest.json file
* @throws FileNotFoundError - when file not found
* @throws InvalidManifestError - when file is not a valid json or can not be parsed to TeamsAppManifest
*/
async function loadFromPath(filePath: string): Promise<TeamsAppManifest> {
 const manifest = await fs.readJson(filePath);
 return manifest;
}
