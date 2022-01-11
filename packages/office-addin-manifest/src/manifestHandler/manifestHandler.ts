import { ManifestInfo } from "../manifestInfo";
import { Xml } from "./manifestHandlerXml";

export abstract class ManifestHandler {
  /* eslint-disable no-unused-vars */
  abstract parseManifest(xml: Xml): ManifestInfo;
  abstract modifyManifestXml(manifestPath: string, guid?: string, displayName?: string): Promise<Xml>;
  abstract parseXmlAsync(xmlString: string, manifestPath: string): Promise<Xml>;
  abstract readXmlFromManifestFile(manifestPath: string): Promise<Xml>;
  abstract setModifiedXmlData(xml: any, guid: string | undefined, displayName: string | undefined): void;
  abstract writeManifestData(manifestPath: string, manifestData: any): Promise<void>;
  /* eslint-enable no-unused-vars */
}
