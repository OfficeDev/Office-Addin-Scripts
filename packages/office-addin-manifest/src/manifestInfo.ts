// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import { usageDataObject } from "./defaults";
import { ManifestHandlerJson } from "./manifestHandler/manifestHandlerJson";
import { ManifestHandlerXml } from "./manifestHandler/manifestHandlerXml";
import { isJsonObject, ManifestHandler } from "./manifestHandler/manifestHandler";
import * as util from "util";
import * as fs from "fs";

export class DefaultSettings {
  public sourceLocation?: string;
  public requestedWidth?: string;
  public requestedHeight?: string;
}

export class ManifestInfo {
  public id?: string;
  public allowSnapshot?: string;
  public alternateId?: string;
  public appDomains?: string[];
  public defaultLocale?: string;
  public description?: string;
  public displayName?: string;
  public highResolutionIconUrl?: string;
  public hosts?: string[];
  public iconUrl?: string;
  public officeAppType?: string;
  public permissions?: string;
  public providerName?: string;
  public supportUrl?: string;
  public version?: string;

  public defaultSettings?: DefaultSettings;
}

export namespace OfficeAddinManifest {
  export async function modifyManifestFile(
    manifestPath: string,
    guid?: string,
    displayName?: string
  ): Promise<ManifestInfo> {
    let manifestData: ManifestInfo = {};
    if (manifestPath) {
      if (guid === undefined && displayName === undefined) {
        throw new Error("You need to specify something to change in the manifest.");
      } else {
        try {
          let manifestHandler: ManifestHandler;
          const fileData: string = await readFromManifestFile(manifestPath);
          if (isJsonObject(fileData)) {
            manifestHandler = new ManifestHandlerJson();
          } else {
            manifestHandler = new ManifestHandlerXml();
          }

          manifestData = await manifestHandler.modifyManifest(manifestPath, fileData, guid, displayName);
          await manifestHandler.writeManifestData(manifestPath, manifestData);
          let output = await readManifestFile(manifestPath);
          usageDataObject.reportSuccess("modifyManifestFile()");
          return output;
        } catch (err: any) {
          usageDataObject.reportException("modifyManifestFile()", err);
          throw err;
        }
      }
    } else {
      throw new Error(`Please provide the path to the manifest file.`);
    }
  }

  export async function readManifestFile(manifestPath: string): Promise<ManifestInfo> {
    if (manifestPath) {
      let manifestHandler: ManifestHandler;
      const fileData: string = await readFromManifestFile(manifestPath);
      if (isJsonObject(fileData)) {
        manifestHandler = new ManifestHandlerJson();
      } else {
        manifestHandler = new ManifestHandlerXml();
      }

      const manifest: ManifestInfo = await manifestHandler.parseManifest(manifestPath, fileData);
      return manifest;
    } else {
      throw new Error(`Please provide the path to the manifest file.`);
    }
  }
}

async function readFromManifestFile(manifestPath: string): Promise<string> {
  try {
    const fileData: string = await util.promisify(fs.readFile)(manifestPath, {
      encoding: "utf8",
    });
    return fileData;
  } catch (err) {
    throw new Error(`Unable to read data for manifest file: ${manifestPath}. \n${err}`);
  }
}
