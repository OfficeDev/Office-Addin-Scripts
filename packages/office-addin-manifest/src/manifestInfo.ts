import * as fs from "fs";
import * as util from "util";
import * as xml2js from "xml2js";
import * as xmlMethods from "./xml";
const readFileAsync = util.promisify(fs.readFile);
const uuid = require('uuid/v1');
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

export async function modifyManifestFile(manifestPath: string, guid?: string, displayName?: string): Promise<ManifestInfo> {
  let manifestData: ManifestInfo = {};
  if (manifestPath) {
    if (!guid && !displayName) {
      throw new Error("You need to specify something to change in the manifest.");
    } else {
      manifestData = await modifyManifestXml(manifestPath, guid, displayName);
      await writeManifestData(manifestPath, manifestData);
      return await readManifestFile(manifestPath);
    }
  } else {
    throw new Error(`Please provide the path to the manifest file.`);
  }
}

async function modifyManifestXml(manifestPath: string, guid?: string, displayName?: string): Promise<Xml> {
  try {
    const manifestXml: Xml = await readXmlFromManifestFile(manifestPath);
    setModifiedXmlData(manifestXml.OfficeApp, guid, displayName);
    return manifestXml;
  } catch (err) {
    throw new Error(`Unable to modify xml data for manifest file: ${manifestPath}. \n${err}`);
  }
}

async function parseXmlAsync(xmlString: string, manifestPath: string): Promise<Xml> {
  return new Promise(async function(resolve, reject) {
    xml2js.parseString(xmlString, function(parseError, xml) {
      if (parseError) {
        reject(new Error(`Unable to parse the manifest file: ${manifestPath}. \n${parseError}`));
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

async function readXmlFromManifestFile(manifestPath: string): Promise<Xml> {
  const fileData: string = await readFileAsync(manifestPath, {encoding: "utf8"});
  const xml = await parseXmlAsync(fileData, manifestPath);
  return xml;
}

function setModifiedXmlData(xml: any, guid: string | undefined, displayName: string | undefined) {
  if (guid) {
    if (guid === "random") {
      guid = uuid();
    }
    xmlMethods.setXmlElementValue(xml, "Id", guid);
  }

  if (displayName) {
    xmlMethods.setXmlElementAttributeValue(xml, "DisplayName", displayName);
  }
}

async function writeManifestData(manifestPath: string, manifestData: any): Promise<void> {
  let xml: Xml;

  try {
    // Generate xml for the manifest data.
    const builder = new xml2js.Builder();
    xml = builder.buildObject(manifestData);
  } catch (err) {
    throw new Error(`Unable to generate xml for the manifest.\n${err}`);
  }

  try {
    // Write the xml back to the manifest file.
    await writeFileAsync(manifestPath, xml);
    } catch (err) {
      throw new Error(`Unable to write to file. ${manifestPath} \n${err}`);
    }
}
