import * as fs from "fs";
import * as util from "util";
import * as xml2js from "xml2js";
import * as xmlMethods from "./xml";
const parseStringAsync = util.promisify(xml2js.parseString);
const readFileAsync = util.promisify(fs.readFile);
const writeFileAsync = util.promisify(fs.writeFile);
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

async function readXmlFromManifestFile(manifestPath: string): Promise<Xml> {
  const fileData: string = await readFileAsync(manifestPath, {encoding: "utf8"});
  const xml = await parseXmlAsync(fileData, manifestPath);
  return xml;
}

async function parseXmlAsync(xmlString: string, manifestPath: string): Promise<Xml> {
  return new Promise(async function(resolve, reject) {
    xml2js.parseString(xmlString, function(parseError, xml) {
      if (parseError) {
        reject(`Unable to parse the xml for manifest file: ${manifestPath}. \n${parseError}`);
      } else {
        resolve(xml);
      }
    });
  });
}

export async function readManifestFile(manifestPath: string): Promise<ManifestInfo> {
  if (manifestPath) {
    const xml = await readXmlFromManifestFile(manifestPath);
    const manifest: ManifestInfo = parseManifest(xml);
    return manifest;
  } else {
    throw new Error(`Please provide the path to the manifest file.`);
  }
}

export async function modifyManifestFile(manifestPath: string, guid?: string, displayName?: string): Promise<ManifestInfo> {
  let manifestData: ManifestInfo = {};
  if (manifestPath) {
    if (!guid && !displayName) {
      throw new Error("You need to specify something to change in the manifest.");
    } else {
      manifestData = await modifyManifestXml(manifestPath, guid, displayName);
      await writeModifiedManifestData(manifestPath, manifestData);
      return await readManifestFile(manifestPath);
    }
  } else {
    throw new Error(`Please provide the path to the manifest file.`);
  }
}

async function modifyManifestXml(manifestPath: string, guid?: string, displayName?: string): Promise<Xml> {
  try {
    const manifestXml: Xml = await readXmlFromManifestFile(manifestPath);
    xmlMethods.setModifiedXmlData(manifestXml.OfficeApp, guid, displayName);
    return manifestXml;
  } catch (err) { throw new Error(`Unable to modify xml data for manifest file: ${manifestPath} \n${err}`); }
}

async function writeModifiedManifestData(manifestPath: string, manifestData: any): Promise<void> {
  try {
    // Regenerate xml from modified manifest data.
    const builder = new xml2js.Builder();
    const xml: Xml = builder.buildObject(manifestData);

    // Write modified xml back to the manifest.
    await writeFileAsync(manifestPath, xml);
  } catch (err) { throw new Error(`Unable to write to file. ${manifestPath} \n${err}`); }
}
