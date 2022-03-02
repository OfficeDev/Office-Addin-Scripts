import * as fs from "fs";
import * as fsExtra from "fs-extra";
import * as jszip from "jszip";
import * as path from "path";
import { ManifestUtil, TeamsAppManifest } from "@microsoft/teams-manifest";

/* global console */

export async function exportMetadataPackage(output: string = "", manifest: string = "manifest.json"): Promise<string> {
  const manifestPath: string = path.resolve(manifest);

  const zip: jszip = await createZip(manifestPath);

  if (output === "") {
    output = path.join(path.dirname(manifestPath), "manifest.zip");
  }
  await saveZip(zip, output);

  return Promise.resolve(output);
}

async function createZip(manifestPath: string = "manifest.json"): Promise<jszip> {
  const zip = new jszip();

  if (fs.existsSync(manifestPath)) {
    const manifestData = fs.readFileSync(manifestPath);
    zip.file("manifest.json", manifestData);
  } else {
    throw new Error(`The file '${manifestPath}' does not exist`);
  }

  // Need json manifest manipulation before this can work.
  const manifest: TeamsAppManifest = await ManifestUtil.loadFromPath(manifestPath);
  addIconFile(manifest.icons?.color, zip);
  addIconFile(manifest.icons?.outline, zip);

  return Promise.resolve(zip);
}

function addIconFile(iconPath: string, zip: jszip) {
  if (iconPath && !iconPath.startsWith("https://")) {
    const filePath: string = path.resolve(iconPath);
    if (fs.existsSync(filePath)) {
      zip.file(iconPath, fs.readFileSync(filePath));
    }
  }
}

async function saveZip(zip: jszip, outputPath: string): Promise<void> {
  outputPath = path.resolve(outputPath);

  fsExtra.ensureDirSync(path.dirname(outputPath));
  return new Promise<void>((fulfill) =>
    zip
      .generateNodeStream({ type: "nodebuffer", streamFiles: true })
      .pipe(fs.createWriteStream(outputPath))
      .on("finish", function () {
        // JSZip generates a readable stream with a "end" event,
        // but is piped here in a writable stream which emits a "finish" event.
        console.log(`Manifest package saved to ${outputPath}`);
        fulfill();
      })
  );
}
