#!/usr/bin/env node

// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as commander from 'commander';
import * as fs from 'fs';
import * as xml2js from 'xml2js';

export interface ManifestInfo {
    id?: string;
    name?: string;
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

function parseManifest(xml: any): ManifestInfo {    
    const manifest: ManifestInfo = { }; 
    const officeApp = xml['OfficeApp'];

    manifest.id = officeApp['Id'][0];

    return manifest;
}
  
commander
    .command('info [path]', 'Display manifest info')
    .action(async function (_command: string, path: string) {
        const manifest = await readManifestFile(path);
        console.log(`Manifest: ${path}`);
        console.log(`  Id: ${manifest.id}`)
    });

commander.parse(process.argv);
