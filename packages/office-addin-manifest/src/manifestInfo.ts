import * as fs from "fs";
import * as xml2js from "xml2js";
import * as xmlMethods from './xml'
const uuid = require('uuid/v1');

export class ManifestInfo {
    public id?: string;
    public defaultLocale?: string;
    public description?: string;
    public displayName?: string;
    public officeAppType?: string;
    public providerName?: string;
    public version?: string;
  }
  
  function parseManifest(xml: any): ManifestInfo {
    const manifest: ManifestInfo = { };
    const officeApp = xml.OfficeApp;
  
    manifest.id = xmlMethods.xmlElementValue(officeApp, "Id");
    manifest.officeAppType = xmlMethods.xmlAttributeValue(officeApp, "xsi:type");
    manifest.defaultLocale = xmlMethods.xmlElementValue(officeApp, "DefaultLocale");
    manifest.description = xmlMethods.xmlElementAttributeValue(officeApp, "Description");
    manifest.displayName = xmlMethods.xmlElementAttributeValue(officeApp, "DisplayName");
    manifest.providerName = xmlMethods.xmlElementValue(officeApp, "ProviderName");
    manifest.version = xmlMethods.xmlElementValue(officeApp, "Version");
  
    return manifest;
  }
   
  export function readManifestFile(manifestPath: string): Promise<ManifestInfo> {
    return new Promise(async function(resolve, reject) {
      if (manifestPath) {
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
      } else {
        reject(`Please provide the path to the manifest file.`);
      }
    });
  }
  
  export function personalizeManifestFile(manifestPath: string, guid?: string, displayName?: string): Promise<void> {  
    return new Promise(async function(resolve, reject) {
      if (manifestPath) {
        try {
          await fs.readFile(manifestPath, function(readError, fileData) {
            if (readError) {
              reject(`Unable to read the manifest file: ${manifestPath}. \n${readError}`);
            } else {
              // tslint:disable-next-line:only-arrow-functions
                xml2js.parseString(fileData, function(parseError, manifestData) {
                if (parseError) {
                  reject(`Unable to parse the manifest file: ${manifestPath}. \n${parseError}`);
                } else {
                  try {
                    // set the guid in the xml
                    if (guid){
                      xmlMethods.setXmlElementValue(manifestData, "Id", guid);
                    }
                    if (displayName){
                      xmlMethods.setElementAttributeValue(manifestData, "DisplayName", displayName); 
                    }
                    // Regenerate xml from manifestData and write xml back to the manifest
                    let builder = new xml2js.Builder();
                    let xml = builder.buildObject(manifestData);
                    const util = require('util');
                    const fs_writeFile = util.promisify(fs.writeFile);
                    fs_writeFile(manifestPath, xml, function(err: string) {
                      if(err) {
                          return console.log(err);
                      }                      
                  });
                 
                  resolve();
                    // resolve (personalizedManifest);                        
                  } catch (err) {
                    reject(`Unable to write to the manifest file: ${manifestPath}. \n${err}`);
                  }
                }
              });                                      
            }
          });
        } catch (err) {
          return reject(`Unable to read the manifest file: ${manifestPath}. \n${err}`);
        }
      } else {
        reject(`Please provide the path to the manifest file.`);
      }
    });
  }