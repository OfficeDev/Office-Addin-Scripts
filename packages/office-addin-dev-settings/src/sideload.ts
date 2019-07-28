// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as fs from "fs";
import * as jszip from "jszip";
import {
  AddInType,
  getAddInTypeForManifestOfficeAppType,
  getOfficeAppName,
  getOfficeAppsForManifestHosts,
  ManifestInfo,
  OfficeApp,
  readManifestFile,
} from "office-addin-manifest";
import open = require("open");
import * as os from "os";
import * as path from "path";
import * as util from "util";
import { registerAddIn } from "./dev-settings-mac";
import { chooseOfficeApp } from "./prompt";

const readFileAsync = util.promisify(fs.readFile);

function getTemplatePath(
  app: OfficeApp,
  addInType: AddInType,
): string | undefined {
  switch (app) {
    case OfficeApp.Excel:
      switch (addInType) {
        case AddInType.Content:
          return path.resolve(__dirname, "../templates/ExcelWorkbookWithContent.xlsx");
        case AddInType.TaskPane:
          return path.resolve(__dirname, "../templates/ExcelWorkbookWithTaskPane.xlsx");
      }
      break;
    case OfficeApp.PowerPoint:
      switch (addInType) {
        case AddInType.Content:
          return path.resolve(__dirname, "../templates/PowerPointPresentationWithContent.pptx");
        case AddInType.TaskPane:
          return path.resolve(__dirname, "../templates/PowerPointPresentationWithTaskPane.pptx");
      }
      break;
    case OfficeApp.Word:
      switch (addInType) {
        case AddInType.TaskPane:
          return path.resolve(__dirname, "../templates/WordDocumentWithTaskPane.docx");
      }
      break;
  }
}

function getWebExtensionPath(
  app: OfficeApp,
  addInType: AddInType,
): string | undefined {
  switch (app) {
    case OfficeApp.Excel:
      return "xl/webextensions/webextension.xml";
    case OfficeApp.PowerPoint:
      switch (addInType) {
        case AddInType.Content:
          return "ppt/slides/udata/data.xml";
        case AddInType.TaskPane:
          return "ppt/webextensions/webextension.xml";
      }
      break;
    case OfficeApp.Word:
      return "word/webextensions/webextension.xml";
  }
}

/**
 * makePathUnique
 * @description Given a file path, returns a unique file path where the file doesn't exist by
 * appending a period and a numeric suffix, starting from 2.
 * @param tryToDelete If true, first try to delete the file if it exists.
 */
function makePathUnique(originalPath: string, tryToDelete: boolean = false): string {
  let currentPath = originalPath;
  let parsedPath = null;
  let suffix = 1;

  while (fs.existsSync(currentPath)) {
    let deleted: boolean = false;

    if (tryToDelete) {
      try {
        fs.unlinkSync(currentPath);
        deleted = true;
      } catch (err) {
        // no error (file is in use)
      }
    }

    if (!deleted) {
      ++suffix;

      if (parsedPath == null) {
        parsedPath = path.parse(originalPath);
      }

      currentPath = path.join(
        parsedPath.dir,
        `${parsedPath.name}.${suffix}${parsedPath.ext}`,
      );
    }
  }

  return currentPath;
}

async function generateSideloadFile(app: OfficeApp, manifest: ManifestInfo): Promise<string> {
  if (!manifest.id) {
    throw new Error("The manifest does not contain the id for the add-in.");
  }

  if (!manifest.officeAppType) {
    throw new Error("The manifest does not contain the OfficeApp xsi:type.");
  }

  if (!manifest.version) {
    throw new Error(
      "The manifest does not contain the version for the add-in.",
    );
  }

  const addInType = getAddInTypeForManifestOfficeAppType(manifest.officeAppType);

  if (!addInType) {
    throw new Error("The manifest contains an unsupported OfficeApp xsi:type.");
  }

  const templatePath = getTemplatePath(app, addInType);

  if (!templatePath) {
    throw new Error("Sideload is not supported.");
  }

  const templateBuffer = await readFileAsync(templatePath);
  const zip = await jszip.loadAsync(templateBuffer);
  const webExtensionPath = getWebExtensionPath(app, addInType);

  if (!webExtensionPath) {
    throw new Error("Don't know the webextension path.");
  }

  const appName = getOfficeAppName(app);
  const extension = path.extname(templatePath);
  const pathToWrite = makePathUnique(
    path.join(os.tmpdir(), `${appName} add-in ${manifest.id}${extension}`),
    true,
  );

  // replace the placeholder id and version
  const webExtensionXml = (await zip.file(webExtensionPath).async("text"))
    .replace(/00000000-0000-0000-0000-000000000000/g, manifest.id)
    .replace(/1.0.0.0/g, manifest.version);
  zip.file(webExtensionPath, webExtensionXml);

  // Write the file
  await new Promise((resolve, reject) => {
    zip
      .generateNodeStream({ type: "nodebuffer", streamFiles: true })
      .pipe(fs.createWriteStream(pathToWrite))
      .on("error", reject)
      .on("finish", resolve);
  });

  return pathToWrite;
}

export async function sideloadAddIn(
  manifestPath: string,
  app?: OfficeApp,
  canPrompt: boolean = false
): Promise<void> {
  const manifest = await readManifestFile(manifestPath);

  if (!app) {
    const appsInManifest = getOfficeAppsForManifestHosts(manifest.hosts);

    if (appsInManifest.length === 1) {
      app = appsInManifest[0];
    } else if (appsInManifest.length === 0) {
      throw new Error("Manifest does not specify Hosts.");
    } else if (canPrompt) {
      app = await chooseOfficeApp(appsInManifest);
    } else {
      throw new Error("Please specify the Office app to sideload.");
    }
  }

  await registerAddIn(manifestPath);

  if (app) {
    const sideloadFile = await generateSideloadFile(app, manifest);

    await open(sideloadFile, { wait: false });
  }
}
