import * as fs from "fs";
import * as fsExtra from "fs-extra";
import * as AdmZip from "adm-zip";
import * as path from "path";
import { ManifestUtil, TeamsAppManifest } from "@microsoft/teams-manifest";

/* global console */

export async function exportMetadataPackage(output: string = "", manifest: string = "manifest.json"): Promise<string> {
  const manifestPath: string = path.resolve(manifest);

  const zip: AdmZip = await createZip(manifestPath);

  if (output === "") {
    output = path.join(path.dirname(manifestPath), "manifest.zip");
  }
  await saveZip(zip, output);

  return Promise.resolve(output);
}

async function createZip(manifestPath: string = "manifest.json"): Promise<AdmZip> {
  const zip: AdmZip = new AdmZip();

  if (fs.existsSync(manifestPath)) {
    zip.addLocalFile(manifestPath, "", "manifest.json");
  } else {
    throw new Error(`The file '${manifestPath}' does not exist`);
  }

  const manifest: TeamsAppManifest = await ManifestUtil.loadFromPath(manifestPath);
  addIconFile(manifest.icons?.color, zip);
  addIconFile(manifest.icons?.outline, zip);

  return Promise.resolve(zip);
}

function addIconFile(iconPath: string, zip: AdmZip) {
  if (iconPath && !iconPath.startsWith("https://")) {
    const filePath: string = path.resolve(iconPath);
    const fileDir: string = path.dirname(iconPath);
    if (fs.existsSync(filePath)) {
      if (!!fileDir && fileDir != ".") {
        zip.addLocalFile(iconPath, fileDir);
      } else {
        zip.addLocalFile(iconPath);
      }
    } else {
      console.log(`Icon File ${filePath} does not exist`);
    }
  }
}

async function saveZip(zip: AdmZip, outputPath: string): Promise<void> {
  outputPath = path.resolve(outputPath);

  fsExtra.ensureDirSync(path.dirname(outputPath));
  const result: Boolean = await zip.writeZipPromise(outputPath);
  if (result) {
    console.log(`Manifest package saved to ${outputPath}`);
  } else {
    throw new Error(`Error writting zip file to ${outputPath}`);
  }
}
