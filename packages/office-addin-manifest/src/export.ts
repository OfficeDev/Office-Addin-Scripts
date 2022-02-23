import * as fs from "fs";
import * as fsExtra from "fs-extra";
import * as jszip from "jszip";
import * as path from "path";

/* global console */

export async function exportMetadataPackage(
  output: string = "",
  manifest: string = "manifest.json",
  assets: string = "assets"
): Promise<string> {
  const manifestPath: string = path.resolve(manifest);

  const zip: jszip = createZip(manifestPath, assets);

  if (output === "") {
    output = path.join(path.dirname(manifestPath), "manifest.zip");
  }
  saveZip(zip, output);

  return output;
}

function createZip(manifestPath: string = "manifest.json", assets: string = "assets"): jszip {
  const zip = new jszip();

  if (fs.existsSync(manifestPath)) {
    const manifestData = fs.readFileSync(manifestPath);
    zip.file("manifest.json", manifestData);
  } else {
    throw new Error(`The file '${manifestPath}' does not exist`);
  }

  if (fs.existsSync(assets)) {
    const files: string[] = fs.readdirSync(assets);
    zip.folder(assets);
    files.forEach((element) => {
      const filePath = path.join(assets, element);
      const fileData = fs.readFileSync(filePath);
      zip.file(filePath, fileData);
    });
  } else {
    throw new Error("Need folder of assets referenced by manifest file");
  }

  return zip;
}

async function saveZip(zip: jszip, outputPath: string) {
  outputPath = path.resolve(outputPath);

  fsExtra.ensureDirSync(path.dirname(outputPath));
  await new Promise(() =>
    zip
      .generateNodeStream({ type: "nodebuffer", streamFiles: true })
      .pipe(fs.createWriteStream(outputPath))
      .on("finish", function () {
        // JSZip generates a readable stream with a "end" event,
        // but is piped here in a writable stream which emits a "finish" event.
        console.log(`Manifest package saved to ${outputPath}`);
      })
  );
}

