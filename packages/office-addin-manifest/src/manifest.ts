#!/usr/bin/env node

// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as commander from "commander";
import * as fs from "fs";
import * as xml2js from "xml2js";

export class ManifestInfo {
    public id?: string;
    public defaultLocale?: string;
    public description?: string;
    public displayName?: string;
    public officeAppType?: string;
    public providerName?: string;
    public version?: string;
}

export function readManifestFile(manifestPath: string): Promise<ManifestInfo> {
    return new Promise(async function(resolve, reject) {
        try {
           fs.readFile(manifestPath, function(readError, fileData) {
                if (readError) {
                    reject(`Unable to read the manifest file: ${manifestPath}. \n${readError}`);
                } else {
                    // tslint:disable-next-line:only-arrow-functions
                    xml2js.parseString(fileData, function(parseError, result) {
                        if (parseError) {
                            reject(`Unable to parse the manifest file: ${manifestPath}. \n${parseError}`);
                        } else {
                            try {
                                const manifest: ManifestInfo = parseManifest(result);
                                resolve (manifest);
                            } catch (err) {
                                reject(`Unable to parse the manifest file: ${manifestPath}. \n${err}`);
                            }
                        }
                    });
                }
            });
        } catch (err) {
          return reject(`Unable to read the manifest file: ${manifestPath}. \n${err}`);
        }
    });
}

function xmlAttributeValue(xml: any, name: string): string | undefined {
    try {
        return xml.$[name];
    } catch (err) {
        // console.error(`Unable to get xml attribute value "${name}". ${err}`);
    }
}

function xmlElementAttributeValue(xml: any, elementName: string, attributeName: string = "DefaultValue"): string | undefined {
    const element = xmlElementValue(xml, elementName);
    if (element) {
        return xmlAttributeValue(element, attributeName);
    }
}

function xmlElementValue(xml: any, name: string): string | undefined {
    try {
        const element = xml[name];

        if (element) {
            return element[0];
        }
    } catch (err) {
        // console.error(`Unable to get xml element value "${name}". ${err}`);
    }
}

function parseManifest(xml: any): ManifestInfo {
    const manifest: ManifestInfo = { };
    const officeApp = xml.OfficeApp;

    manifest.id = xmlElementValue(officeApp, "Id");
    manifest.officeAppType = xmlAttributeValue(officeApp, "xsi:type");
    manifest.defaultLocale = xmlElementValue(officeApp, "DefaultLocale");
    manifest.description = xmlElementAttributeValue(officeApp, "Description");
    manifest.displayName = xmlElementAttributeValue(officeApp, "DisplayName");
    manifest.providerName = xmlElementValue(officeApp, "ProviderName");
    manifest.version = xmlElementValue(officeApp, "Version");

    return manifest;
}

async function infoCommandAction(path: string) {
    try {
        const manifest = await readManifestFile(path);

        console.log(`Manifest: ${path}`);
        console.log(`  Id: ${manifest.id || ""}`);
        console.log(`  Name: ${manifest.displayName || ""}`);
        console.log(`  Provider: ${manifest.providerName || ""}`);
        console.log(`  Type: ${manifest.officeAppType || ""}`);
        console.log(`  Version: ${manifest.version || ""}`);
        console.log(`  Default Locale: ${manifest.defaultLocale || ""}`);
        console.log(`  Description: ${manifest.description || ""}`);
    } catch (err) {
        console.error(`Error: ${err}`);
    }
}

commander
    .command("info [path]")
    .action(infoCommandAction);

commander.parse(process.argv);
