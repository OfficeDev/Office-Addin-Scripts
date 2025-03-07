import fs from "fs";
import fsExtra from "fs-extra";
import AdmZip from "adm-zip";
import path from "path";
import { DeclarativeCopilotManifestSchema } from "@microsoft/teams-manifest";
import { DevPreviewSchema } from "./devPreviewManifest";

/* global console */

export async function exportMetadataPackage(
  output: string = "",
  manifest: string = "manifest.json"
): Promise<string> {
  const zip: AdmZip = await createZip(manifest);

  if (output === "") {
    output = path.join(path.dirname(path.resolve(manifest)), "manifest.zip");
  }
  await saveZip(zip, output);

  return Promise.resolve(output);
}

function readJsonFileSync<JsonType>(filePath: string): JsonType {
  const jsonText = fs.readFileSync(filePath, "utf8");
  const jsonObject: JsonType = JSON.parse(jsonText);
  return jsonObject;
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

  const manifest: DevPreviewSchema = readJsonFileSync(manifestPath);
  const agents = manifest?.copilotAgents?.declarativeAgents;

  if (agents) {
    agents.forEach((agent, agentIndex) => {
      const file: string = agent?.file;
      const filePath = path.join(manifestDir, file);

      if (!fs.existsSync(filePath)) {
        throw new Error(
          `copilotAgents.declarativeAgents[${agentIndex}].file does not exist:  ${filePath}`
        );
      }

      zip.addLocalFile(filePath, "", file);
      const agentJson: DeclarativeCopilotManifestSchema = readJsonFileSync(filePath);

      agentJson?.actions?.forEach((action, actionIndex) => {
        if (action?.file) {
          const actionFilePath = path.join(manifestDir, action.file);
          if (!fs.existsSync(actionFilePath)) {
            throw new Error(`actions[${actionIndex}].file does not exist: ${actionFilePath}`);
          }

          zip.addLocalFile(actionFilePath, "", action.file);
        }
      });
    });
  }

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
