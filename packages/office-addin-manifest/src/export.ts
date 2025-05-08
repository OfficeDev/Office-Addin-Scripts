import fs from "fs";
import fsExtra from "fs-extra";
import AdmZip from "adm-zip";
import path from "path";
import {
  AppManifestUtils,
  DeclarativeAgentManifest,
  TeamsManifestVDevPreview,
} from "@microsoft/app-manifest";

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

async function createZip(manifestPath: string): Promise<AdmZip> {
  const absolutePath: string = path.resolve(manifestPath);
  const manifestDir: string = path.dirname(absolutePath);
  const zip: AdmZip = new AdmZip();

  if (fs.existsSync(manifestPath)) {
    zip.addLocalFile(manifestPath, "", "manifest.json");
  } else {
    throw new Error(`The file '${manifestPath}' does not exist`);
  }

  const manifest: TeamsManifestVDevPreview = (await AppManifestUtils.readTeamsManifest(
    manifestPath
  )) as TeamsManifestVDevPreview;

  // Add icons
  addZipFile(manifest.icons?.color, manifestDir, zip);
  addZipFile(manifest.icons?.outline, manifestDir, zip);

  // Add localization files
  const languages = manifest.localizationInfo?.additionalLanguages;
  if (languages) {
    languages.forEach((language) => {
      addZipFile(language?.file, manifestDir, zip);
    });
  }

  // Add Declarative Copilot Agents
  const agents = manifest?.copilotAgents?.declarativeAgents;
  if (agents) {
    for (const agent of agents) {
      const agentFile: string = agent?.file;
      addZipFile(agentFile, manifestDir, zip);

      const agentRelDir: string = path.dirname(agentFile);
      const agentManifest: DeclarativeAgentManifest =
        await AppManifestUtils.readDeclarativeAgentManifest(path.join(manifestDir, agentFile));
      agentManifest?.actions?.forEach((action) => {
        if (action?.file) {
          addZipFile(path.join(agentRelDir, action.file), manifestDir, zip);
        }
      });
    }
  }

  return Promise.resolve(zip);
}

function addZipFile(filePath: string, baseDir: string, zip: AdmZip) {
  if (filePath && !filePath.startsWith("https://")) {
    const fullPath: string = path.join(baseDir, filePath);
    const fileDir: string = path.dirname(filePath);
    if (fs.existsSync(fullPath)) {
      zip.addLocalFile(fullPath, fileDir === "." ? "" : fileDir);
    } else {
      throw new Error(`File to zip "${fullPath}' does not exist`);
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
