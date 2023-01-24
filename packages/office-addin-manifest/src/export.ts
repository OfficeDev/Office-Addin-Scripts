import * as fs from "fs";
import * as fsExtra from "fs-extra";
import * as AdmZip from "adm-zip";
import * as path from "path";
import { ManifestUtil, TeamsAppManifest } from "@microsoft/teams-manifest";

/* global console */

export async function exportMetadataPackage(output: string = "", manifest: string = "manifest.json"): Promise<string> {
  const zip: AdmZip = await createZip(manifest);

  if (output === "") {
    output = path.join(path.dirname(path.resolve(manifest)), "manifest.zip");
  }
  await saveZip(zip, output);

  return Promise.resolve(output);
}

async function createZip(manifestPath: string): Promise<AdmZip> {
  const absolutePath: string = path.resolve(manifestPath);
  const manifestDir: string = path.dirname(absolutePath);
  const zip: AdmZip = new AdmZip();

  if (fs.existsSync(manifestPath)) {
    zip.addLocalFile(manifestPath, "", "manifest.json");
  } else {
    throw new Error(`The file '${manifestPath}' does not exist`);
  }

  const manifest: TeamsAppManifest = await ManifestUtil.loadFromPath(manifestPath);
  addIconFile(manifest.icons?.color, manifestDir, zip);
  addIconFile(manifest.icons?.outline, manifestDir, zip);

  return Promise.resolve(zip);
}

function addIconFile(iconPath: string, manifestDir: string, zip: AdmZip) {
  if (iconPath && !iconPath.startsWith("https://")) {
    const filePath: string = path.join(manifestDir, iconPath);
    const iconDir: string = path.dirname(iconPath);
    if (fs.existsSync(filePath)) {
      zip.addLocalFile(filePath, iconDir === "." ? "" : iconDir);
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
