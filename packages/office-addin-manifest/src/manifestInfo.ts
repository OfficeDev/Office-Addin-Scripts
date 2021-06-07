// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as fs from "fs";
import * as util from "util";
import { v4 as uuidv4 } from "uuid";
import * as xml2js from "xml2js";
import * as xmlMethods from "./xml";
import { usageDataObject } from "./defaults";
const readFileAsync = util.promisify(fs.readFile);
const writeFileAsync = util.promisify(fs.writeFile);
type Xml = xmlMethods.Xml;

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

function parseManifest(xml: Xml): ManifestInfo {
  const manifest: ManifestInfo = new ManifestInfo();
  const officeApp: Xml = xml.OfficeApp;

  if (officeApp) {
    const defaultSettingsXml: Xml = xmlMethods.getXmlElement(officeApp, "DefaultSettings");

    manifest.id = xmlMethods.getXmlElementValue(officeApp, "Id");
    manifest.allowSnapshot = xmlMethods.getXmlElementValue(officeApp, "AllowSnapshot");
    manifest.alternateId = xmlMethods.getXmlElementValue(officeApp, "AlternateId");
    manifest.appDomains = xmlMethods.getXmlElementsValue(officeApp, "AppDomains", "AppDomain");
    manifest.defaultLocale = xmlMethods.getXmlElementValue(officeApp, "DefaultLocale");
    manifest.description = xmlMethods.getXmlElementAttributeValue(officeApp, "Description");
    manifest.displayName = xmlMethods.getXmlElementAttributeValue(officeApp, "DisplayName");
    manifest.highResolutionIconUrl = xmlMethods.getXmlElementAttributeValue(officeApp, "HighResolutionIconUrl");
    manifest.hosts = xmlMethods.getXmlElementsAttributeValue(officeApp, "Hosts", "Host", "Name");
    manifest.iconUrl = xmlMethods.getXmlElementAttributeValue(officeApp, "IconUrl");
    manifest.officeAppType = xmlMethods.getXmlAttributeValue(officeApp, "xsi:type");
    manifest.permissions = xmlMethods.getXmlElementValue(officeApp, "Permissions");
    manifest.providerName = xmlMethods.getXmlElementValue(officeApp, "ProviderName");
    manifest.supportUrl = xmlMethods.getXmlElementAttributeValue(officeApp, "SupportUrl");
    manifest.version = xmlMethods.getXmlElementValue(officeApp, "Version");

    if (defaultSettingsXml) {
      const defaultSettings: DefaultSettings = new DefaultSettings();

      defaultSettings.requestedHeight = xmlMethods.getXmlElementValue(defaultSettingsXml, "RequestedHeight");
      defaultSettings.requestedWidth = xmlMethods.getXmlElementValue(defaultSettingsXml, "RequestedWidth");
      defaultSettings.sourceLocation = xmlMethods.getXmlElementAttributeValue(defaultSettingsXml, "SourceLocation");

      manifest.defaultSettings = defaultSettings;
    }
  }

  return manifest;
}

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
      } catch (err) {
        usageDataObject.reportException("modifyManifestFile()", err);
        throw err;
      }
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
  const fileData: string = await readFileAsync(manifestPath, {
    encoding: "utf8",
  });
  const xml = await parseXmlAsync(fileData, manifestPath);
  return xml;
}

function setModifiedXmlData(xml: any, guid: string | undefined, displayName: string | undefined): void {
  if (typeof guid !== "undefined") {
    if (!guid || guid === "random") {
      guid = uuidv4();
    }
    xmlMethods.setXmlElementValue(xml, "Id", guid);
  }

  if (typeof displayName !== "undefined") {
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
