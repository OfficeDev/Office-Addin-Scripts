import * as fs from "fs";
import * as fsExtra from "fs-extra";
import * as jszip from "jszip";
import * as path from "path";

export function exportMetadataPackage(
  output: string = "package/manifest.zip",
  manifest: string = "manifest.json",
  assets: string = "assets"
) {
  const zip = new jszip();

  if (fs.existsSync(manifest)) {
    const manifestData = fs.readFileSync(manifest);
    zip.file("manifest.json", manifestData);
  } else {
    throw new Error(`The file ${manifest} does not exist`);
  }

  if (fs.existsSync(assets)) {
    const files: string[] = fs.readdirSync(assets);
    zip.folder(assets);
    files.forEach((element) => {
      const filePath = assets + "/" + element;
      const fileData = fs.readFileSync(filePath);
      zip.file(filePath, fileData);
    });
  } else {
    throw new Error("Need folder of assets referenced by manifest file");
  }

  fsExtra.ensureDirSync(path.dirname(output));
  zip.generateNodeStream({ type: "nodebuffer", streamFiles: true }).pipe(fs.createWriteStream(output));
}
