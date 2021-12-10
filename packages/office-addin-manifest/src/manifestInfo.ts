// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as fs from "fs";
import * as util from "util";
import { v4 as uuidv4 } from "uuid";
import * as xml2js from "xml2js";
import { ManifestXML, Xml } from "./parser/xml";
import { usageDataObject } from "./defaults";
const readFileAsync = util.promisify(fs.readFile);
const writeFileAsync = util.promisify(fs.writeFile);

class DefaultSettings {
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
          manifestData = await modifyManifestXml(manifestPath, guid, displayName);
          await writeManifestData(manifestPath, manifestData);
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
      const xml = await readXmlFromManifestFile(manifestPath);
      const manifest: ManifestInfo = parseManifest(xml);
      return manifest;
    } else {
      throw new Error(`Please provide the path to the manifest file.`);
    }
  }
}

function parseManifest(xml: Xml): ManifestInfo {
  const manifest: ManifestInfo = new ManifestInfo();
  const officeApp: Xml = xml.OfficeApp;
  const xmlMethods: ManifestXML = new ManifestXML(officeApp);

  if (officeApp) {
    manifest.id = xmlMethods.getElementValue("Id");
    manifest.allowSnapshot = xmlMethods.getElementValue("AllowSnapshot");
    manifest.alternateId = xmlMethods.getElementValue("AlternateId");
    manifest.appDomains = xmlMethods.getElementsValue("AppDomains", "AppDomain");
    manifest.defaultLocale = xmlMethods.getElementValue("DefaultLocale");
    manifest.description = xmlMethods.getElementAttributeValue("Description");
    manifest.displayName = xmlMethods.getElementAttributeValue("DisplayName");
    manifest.highResolutionIconUrl = xmlMethods.getElementAttributeValue("HighResolutionIconUrl");
    manifest.hosts = xmlMethods.getElementsAttributeValue("Hosts", "Host", "Name");
    manifest.iconUrl = xmlMethods.getElementAttributeValue("IconUrl");
    manifest.officeAppType = xmlMethods.getAttributeValue("xsi:type");
    manifest.permissions = xmlMethods.getElementValue("Permissions");
    manifest.providerName = xmlMethods.getElementValue("ProviderName");
    manifest.supportUrl = xmlMethods.getElementAttributeValue("SupportUrl");
    manifest.version = xmlMethods.getElementValue("Version");

    const defaultSettingsXml: Xml = xmlMethods.getElement("DefaultSettings");
    if (defaultSettingsXml) {
      const defaultSettingsManifestXml: ManifestXML = new ManifestXML(defaultSettingsXml);
      const defaultSettings: DefaultSettings = new DefaultSettings();

      defaultSettings.requestedHeight = defaultSettingsManifestXml.getElementValue("RequestedHeight");
      defaultSettings.requestedWidth = defaultSettingsManifestXml.getElementValue("RequestedWidth");
      defaultSettings.sourceLocation = defaultSettingsManifestXml.getElementAttributeValue("SourceLocation");

      manifest.defaultSettings = defaultSettings;
    }
  }

  return manifest;
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
  return new Promise(function (resolve, reject) {
    xml2js.parseString(xmlString, function (parseError, xml) {
      if (parseError) {
        reject(new Error(`Unable to parse the manifest file: ${manifestPath}. \n${parseError}`));
      } else {
        resolve(xml);
      }
    });
  });
}

async function readXmlFromManifestFile(manifestPath: string): Promise<Xml> {
  const fileData: string = await readFileAsync(manifestPath, {
    encoding: "utf8",
  });
  const xml = await parseXmlAsync(fileData, manifestPath);
  return xml;
}

function setModifiedXmlData(xml: any, guid: string | undefined, displayName: string | undefined): void {
  const manifestXML: ManifestXML = new ManifestXML(xml);
  if (typeof guid !== "undefined") {
    if (!guid || guid === "random") {
      guid = uuidv4();
    }
    manifestXML.setElementValue("Id", guid);
  }

  if (typeof displayName !== "undefined") {
    manifestXML.setElementAttributeValue("DisplayName", displayName);
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
