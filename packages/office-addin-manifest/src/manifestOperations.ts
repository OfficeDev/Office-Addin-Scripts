import { usageDataObject } from "./defaults";
import { ManifestHandler } from "./manifestHandler/manifestHandler";
import { getManifestHandler } from "./manifestHandler/getManifestHandler";
import * as util from "util";
import * as fs from "fs";
import { ManifestInfo } from "./manifestInfo";

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
          const manifestHandler: ManifestHandler = getManifestHandler(manifestPath);
          const fileData: string = await readFromManifestFile(manifestPath);

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
      const manifestHandler: ManifestHandler = getManifestHandler(manifestPath);
      const fileData: string = await readFromManifestFile(manifestPath);

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
