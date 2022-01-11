// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as fs from "fs";
import * as util from "util";
import { v4 as uuidv4 } from "uuid";
import * as xml2js from "xml2js";
import * as xmlMethods from "../xml";
import { DefaultSettings, ManifestInfo } from "../manifestInfo";
import { ManifestHandler } from "./manifestHandler";
const readFileAsync = util.promisify(fs.readFile);
const writeFileAsync = util.promisify(fs.writeFile);
export type Xml = xmlMethods.Xml;

export class ManifestHandlerXml implements ManifestHandler {
  constructor() {}

  parseManifest(xml: Xml): ManifestInfo {
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

  async modifyManifest(manifestPath: string, guid?: string, displayName?: string): Promise<Xml> {
    try {
      const manifestXml: Xml = await this.readFromManifestFile(manifestPath);
      this.setModifiedXmlData(manifestXml.OfficeApp, guid, displayName);
      return manifestXml;
    } catch (err) {
      throw new Error(`Unable to modify xml data for manifest file: ${manifestPath}. \n${err}`);
    }
  }

  async parseXmlAsync(xmlString: string, manifestPath: string): Promise<Xml> {
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

  async readFromManifestFile(manifestPath: string): Promise<Xml> {
    const fileData: string = await readFileAsync(manifestPath, {
      encoding: "utf8",
    });
    const xml = await this.parseXmlAsync(fileData, manifestPath);
    return xml;
  }

  setModifiedXmlData(xml: any, guid: string | undefined, displayName: string | undefined): void {
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

  async writeManifestData(manifestPath: string, manifestData: any): Promise<void> {
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
}
