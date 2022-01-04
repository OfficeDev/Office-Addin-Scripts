import * as fs from "fs";
import * as fsExtra from "fs-extra";
import * as jszip from "jszip";
import * as path from "path";

export function exportMetadataPackage(
  output: string = "",
  manifest: string = "manifest.json",
  assets: string = "assets"
): string {
  const zip = new jszip();
  const manifestPath: string = path.resolve(manifest);

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

  if (output === "") {
    output = path.join(path.dirname(manifestPath), "manifest.zip");
  } else {
    output = path.resolve(output);
  }

  fsExtra.ensureDirSync(path.dirname(output));
  zip.generateNodeStream({ type: "nodebuffer", streamFiles: true }).pipe(fs.createWriteStream(output));

  return output;
}
