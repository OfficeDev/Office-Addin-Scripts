#!/usr/bin/env node

// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as commander from 'commander';
import * as fs from 'fs';
import * as xml2js from 'xml2js';

export interface ManifestInfo {
    id?: string;
    defaultLocale?: string;
    description?: string;
    displayName?: string;
    officeAppType?: string;
    providerName?: string;
    version?: string;
}

export function readManifestFile(manifestPath: string): Promise<ManifestInfo> {
    return new Promise(async function (resolve, reject) {
        try {
           fs.readFile(manifestPath, function (err, fileData) {  
                if (err) {
                    reject(`Failed to read the manifest file: ${manifestPath}`);          
                } else {
                    xml2js.parseString(fileData, function (err, result) {
                        if (err) {
                            reject(`Failed to parse the manifest file: ${manifestPath}`);
                        } else {
                            try {
                                const manifest: ManifestInfo = parseManifest(result);
                                resolve (manifest);
                            } catch(exception) {
                                reject(`Failed to parse the manifest file: ${manifestPath}`);
                            }
                        }
                    });
                }
            });            
        } catch (err) {
          return reject(['Failed to read the manifest file: ', manifestPath]);
        }
    });
}

function xmlAttributeValue(xml: any, name: string): string {    
    return xml['$'][name];
}

function xmlElementAttributeValue(xml: any, elementName: string, attributeName: string = 'DefaultValue'): string {
    const element = xmlElementValue(xml, elementName);
    return xmlAttributeValue(element, attributeName);
}

function xmlElementValue(xml: any, name: string): string {
    return xml[name][0];
}

function parseManifest(xml: any): ManifestInfo {    
    const manifest: ManifestInfo = { }; 
    const officeApp = xml['OfficeApp'];
    
    manifest.id = xmlElementValue(officeApp, 'Id');
    manifest.officeAppType = xmlAttributeValue(officeApp, 'xsi:type');
    manifest.defaultLocale = xmlElementValue(officeApp, 'DefaultLocale');
    manifest.description = xmlElementAttributeValue(officeApp, 'Description');
    manifest.displayName = xmlElementAttributeValue(officeApp, 'DisplayName');
    manifest.providerName = xmlElementValue(officeApp, 'ProviderName');
    manifest.version = xmlElementValue(officeApp, 'Version');

    return manifest;
}
  
async function infoCommandAction(path: string) {
    const manifest = await readManifestFile(path);

    console.log(`Manifest: ${path}`);
    console.log(`  Id: ${manifest.id}`)
    console.log(`  Name: ${manifest.displayName}`)
    console.log(`  Provider: ${manifest.providerName}`)
    console.log(`  Type: ${manifest.officeAppType}`)
    console.log(`  Version: ${manifest.version}`)
    console.log(`  Default Locale: ${manifest.defaultLocale}`)
    console.log(`  Description: ${manifest.description}`)
}

commander
    .command('info [path]')
    .action(infoCommandAction);

commander.parse(process.argv);
