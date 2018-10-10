import * as fs from "fs";
import * as xml2js from "xml2js";
import * as xmlMethods from "./xml";
type Xml = any;

export class ManifestInfo {
  public id?: string;
  public defaultLocale?: string;
  public description?: string;
  public displayName?: string;
  public officeAppType?: string;
  public providerName?: string;
  public version?: string;
}

function parseManifest(xml: Xml): ManifestInfo {
const manifest: ManifestInfo = { };
const officeApp = xml.OfficeApp;

manifest.id = xmlMethods.getXmlElementValue(officeApp, "Id");
manifest.officeAppType = xmlMethods.getXmlAttributeValue(officeApp, "xsi:type");
manifest.defaultLocale = xmlMethods.getXmlElementValue(officeApp, "DefaultLocale");
manifest.description = xmlMethods.getXmlElementAttributeValue(officeApp, "Description");
manifest.displayName = xmlMethods.getXmlElementAttributeValue(officeApp, "DisplayName");
manifest.providerName = xmlMethods.getXmlElementValue(officeApp, "ProviderName");
manifest.version = xmlMethods.getXmlElementValue(officeApp, "Version");

return manifest;
}

export function readXmlFromManifestFile(manifestPath: string): Promise<Xml> {
  return new Promise(async function(resolve, reject) {
    try {
      fs.readFile(manifestPath, function(readError, fileData) {
        if (readError) {
          reject(`Unable to read the manifest file: ${manifestPath}. \n${readError}`);
        } else {
          // tslint:disable-next-line:only-arrow-functions
          xml2js.parseString(fileData, function(parseError, xml) {
            if (parseError) {
              reject(`Unable to parse the manifest file: ${manifestPath}. \n${parseError}`);
            } else { resolve(xml); }
          });
        }
      });
    } catch (err) {reject(`Unable to read the manifest file: ${manifestPath}. \n${err}`); }
  });
}

export function readManifestFile(manifestPath: string): Promise<ManifestInfo> {
  return new Promise(async function(resolve, reject) {
    if (manifestPath) {
      let result: Xml;
      try {
        result = await readXmlFromManifestFile(manifestPath);
      } catch (err) { reject(err); }

      if (result) {
        try {
          const manifest: ManifestInfo = parseManifest(result);
          resolve (manifest);
        } catch (err) {
          reject(`Unable to parse the manifest file: ${manifestPath}. \n${err}`);
        }
      }
    } else {
      reject(`Please provide the path to the manifest file.`);
    }
  });
}

export async function modifyManifestFile(manifestPath: string, guid?: string, displayName?: string): Promise<any> {
  let manifestData: ManifestInfo = {};
  if (manifestPath) {
    try {
      if (!guid && !displayName) {
        throw new Error("You need to specify something to change in the manifest.");
      } else {
        manifestData = await modifyManifestXml(manifestPath, guid, displayName);
        await writeModifiedManifestData(manifestPath, manifestData);
        return await readManifestFile(manifestPath);
      }
      } catch (err) { return err; }
    }
  return manifestData;
}

export function modifyManifestXml(manifestPath: string, guid?: string, displayName?: string): Promise<any> {
  return new Promise(async function(resolve, reject) {
    try {
      let manifestData;
      try {
        manifestData = await readXmlFromManifestFile(manifestPath);
      } catch { reject(`Unable to read and parse the manifest file: ${manifestPath}.`); }
      // set the guid and displayName in the xml
      xmlMethods.setModifiedXmlData(manifestData.OfficeApp, guid, displayName);
      resolve(manifestData);
    } catch (err) { reject(`Unable to read the manifest file: ${manifestPath}. \n${err}`); }
  });
}

export function writeModifiedManifestData(manifestPath: string, manifestData: any): Promise<void> {
  return new Promise(async function(resolve, reject) {
    // Regenerate xml from manifestData and write xml back to the manifest
    try {
      const builder = new xml2js.Builder();
      const xml = builder.buildObject(manifestData);

      await fs.writeFile(manifestPath, xml, function(err) {
        if (err ) {
            reject(`Unable to write to the manifest file:  ${manifestPath}. \n${err}`);
        } else { resolve(); }
      });
    } catch {reject(`Unable to write to the manifest file:  ${manifestPath}.`); }
  });
}
