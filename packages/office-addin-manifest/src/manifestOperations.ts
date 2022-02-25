// copyright (c) Microsoft Corporation. All rights reserved.
// licensed under the MIT license.

import { usageDataObject } from "./defaults";
import { ManifestHandler } from "./manifestHandler/manifestHandler";
import { ManifestInfo } from "./manifestInfo";
import { ManifestHandlerJson } from "./manifestHandler/manifestHandlerJson";
import { ManifestHandlerXml } from "./manifestHandler/manifestHandlerXml";

export namespace OfficeAddinManifest {
  export async function modifyManifestFile(
    manifestPath: string,
    guid?: string,
    displayName?: string
  ): Promise<ManifestInfo> {
    if (manifestPath) {
      if (guid === undefined && displayName === undefined) {
        throw new Error("You need to specify something to change in the manifest.");
      } else {
        try {
          const manifestHandler: ManifestHandler = await getManifestHandler(manifestPath);
          const manifestData = await manifestHandler.modifyManifest(guid, displayName);
          await manifestHandler.writeManifestData(manifestData);
          const manifestInfo: ManifestInfo = await manifestHandler.parseManifest();
          usageDataObject.reportSuccess("modifyManifestFile()");
          return manifestInfo;
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
      const manifestHandler: ManifestHandler = await getManifestHandler(manifestPath);
      const manifest: ManifestInfo = await manifestHandler.parseManifest();
      return manifest;
    } else {
      throw new Error(`Please provide the path to the manifest file.`);
    }
  }
}

async function getManifestHandler(manifestPath: string): Promise<ManifestHandler> {
  let manifestHandler: ManifestHandler;
  if (manifestPath.endsWith(".json")) {
    manifestHandler = new ManifestHandlerJson(manifestPath);
  } else if (manifestPath.endsWith(".xml")) {
    manifestHandler = new ManifestHandlerXml(manifestPath);
  } else {
    const extension: string = manifestPath.split(".").pop() ?? "<no extension>";
    throw new Error(
      `Manifest operations are not supported in .${extension}.\nThey are only supported in .xml and in .json.`
    );
  }
  await manifestHandler.updateFileData();
  return manifestHandler;
}
