import * as fs from "fs";
import * as xml2js from "xml2js";
import * as xmlMethods from './xml'

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

  export function personalizeManifest(manifestPath: string, guid?: string, displayName?: string): Promise<ManifestInfo> {
    return new Promise(async function(resolve, reject) {
      let manifestData: any = undefined;
      let manifestInfo = undefined;

      if (manifestPath) {
        try {
          if (guid == undefined && displayName == undefined) {
            reject(`Please provide either a guid or displayName parameter.`);
          }
          else {
            try {
              manifestData = await personalizeManifestXml(manifestPath, guid, displayName);
            }
            catch {
              reject('Unable to generate personalized manifest xml.')            
            }

            try {
              await writePersonalizedManifestData(manifestPath, manifestData);
            }
            catch {
              reject(`Unable to write peronalized manifest XML to manifest file: ${manifestPath}`);
            }
            try {
              manifestInfo = await readManifestFile(manifestPath);
              if (manifestInfo) {
                resolve(manifestInfo);
              }
            }
            catch {
              reject(`Unable to read info from manifest file: ${manifestPath}`);
            }
          }
        }
        catch {
          reject(`Please provide the path to the manifest file.`);
        }
      }
    });
  }
  
  export function personalizeManifestXml(manifestPath: string, guid?: string, displayName?: string): Promise<any> {  
    return new Promise(async function(resolve, reject) {
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
                // set the guid and displayName in the xml
                xmlMethods.setPersonalizedXmlData(manifestData.OfficeApp, guid, displayName);
                resolve(manifestData);
              }
            });
          }
        });
      } catch (err) {
        return reject(`Unable to read the manifest file: ${manifestPath}. \n${err}`);
      }
  });
  }

  export function writePersonalizedManifestData(manifestPath: string, manifestData: any) : Promise<void> {
    return new Promise(async function(resolve, reject) {
      // Regenerate xml from manifestData and write xml back to the manifest
      try{
        let builder = new xml2js.Builder();
        let xml = builder.buildObject(manifestData);

        await fs.writeFile(manifestPath, xml, function(err) {
          if(err) {
              reject(`Unable to write to the manifest file:  ${manifestPath}. \n${err}`)
          }
          else{
            resolve();
          }
      });
      }
      catch {
        reject(`Unable to write to the manifest file:  ${manifestPath}.`)
      } 
    });
  }