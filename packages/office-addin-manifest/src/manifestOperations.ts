import { usageDataObject } from "./defaults";
import { ManifestHandler } from "./manifestHandler/manifestHandler";
import { getManifestHandler } from "./manifestHandler/getManifestHandler";
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
          const manifestHandler: ManifestHandler = await getManifestHandler(manifestPath);
          manifestData = await manifestHandler.modifyManifest(guid, displayName);
          await manifestHandler.writeManifestData(manifestData);
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
      const manifestHandler: ManifestHandler = await getManifestHandler(manifestPath);
      const manifest: ManifestInfo = await manifestHandler.parseManifest();
      return manifest;
    } else {
      throw new Error(`Please provide the path to the manifest file.`);
    }
  }
}
