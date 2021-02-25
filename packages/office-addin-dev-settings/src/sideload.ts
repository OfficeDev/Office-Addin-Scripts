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
import { AppType } from "./appType";
import { registerAddIn } from "./dev-settings";
import { startDetachedProcess } from "./process";
import { chooseOfficeApp } from "./prompt";
import * as registry from "./registry";
import { usageDataObject } from './defaults';
import { ExpectedError } from "office-addin-usage-data";

const readFileAsync = util.promisify(fs.readFile);

/**
 * Create an Office document in the temporary files directory
 * which can be opened to launch the Office app and load the add-in.
 * @param app Office app
 * @param manifest Manifest for the add-in.
 * @returns Path to the file.
 */
export async function generateSideloadFile(app: OfficeApp, manifest: ManifestInfo, document?: string): Promise<string> {
  if (!manifest.id) {
    throw new ExpectedError("The manifest does not contain the id for the add-in.");
  }

  if (!manifest.officeAppType) {
    throw new ExpectedError("The manifest does not contain the OfficeApp xsi:type.");
  }

  if (!manifest.version) {
    throw new ExpectedError(
      "The manifest does not contain the version for the add-in.",
    );
  }
  const addInType = getAddInTypeForManifestOfficeAppType(manifest.officeAppType);

  if (!addInType) {
    throw new ExpectedError("The manifest contains an unsupported OfficeApp xsi:type.");
  }

  const templatePath = document && document !== "" ? path.resolve(document) : getTemplatePath(app, addInType);

  if (!templatePath) {
    throw new ExpectedError("Sideload is not supported.");
  }

  const templateBuffer = await readFileAsync(templatePath);
  const zip = await jszip.loadAsync(templateBuffer);
  const webExtensionPath = getWebExtensionPath(app, addInType);

  if (!webExtensionPath) {
    throw new ExpectedError("Don't know the webextension path.");
  }

  const appName = getOfficeAppName(app);
  const extension = path.extname(templatePath);
  const pathToWrite = makePathUnique(
    path.join(os.tmpdir(), `${appName} add-in ${manifest.id}${extension}`),
    true,
  );

  // replace the placeholder id and version
  const zipFile = zip.file(webExtensionPath);
  if (!zipFile) {
    throw new ExpectedError("webextension was not found.")
  }
  const webExtensionXml = (await zipFile.async("text"))
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

/**
 * Create an Office document url with query params which can be opened
 * to register an Office add-in in Office Online.
 * @param manifestPath Path to the manifest file for the Office Add-in.
 * @param documentUrl Office Online document url
 * @param isTest Indicates whether to append test query param to suppress Office Online dialogs.
 * @returns Document url with query params appended.
 */
export async function generateSideloadUrl(manifestFileName: string, manifest: ManifestInfo, documentUrl: string | undefined,  isTest: boolean = false): Promise<string | undefined> {
  const testQueryParam = "&wdaddintest=true";

  if (documentUrl === undefined || documentUrl === "") {
    return undefined;
  }

  if (!manifest.id) {
    throw new ExpectedError("The manifest does not contain the id for the add-in.");
  }

  if (manifest.defaultSettings === undefined || manifest.defaultSettings.sourceLocation === undefined) {
    throw new ExpectedError("The manifest does not contain the SourceLocation for the add-in")
  }

  const sourceLocationUrl: URL = new URL(manifest.defaultSettings.sourceLocation);
  if (sourceLocationUrl.protocol.indexOf("https") === -1) {
    throw new ExpectedError("The SourceLocation in the manifest does not use the HTTPS protocol.");
  }

  if (sourceLocationUrl.host.indexOf("localhost") === -1 && sourceLocationUrl.host.indexOf("127.0.0.1") === -1) {
    throw new ExpectedError("The hostname specified by the SourceLocation in the manifest is not supported for sideload. The hostname should be 'localhost' or 127.0.0.1.");
  }

  let queryParms: string = `&wdaddindevserverport=${sourceLocationUrl.port}&wdaddinmanifestfile=${manifestFileName}&wdaddinmanifestguid=${manifest.id}`;

  if (isTest) {
    queryParms = `${queryParms}${testQueryParam}`;
  }

  return `${documentUrl}${queryParms}`;
}

/**
 * Returns the path to the document used as a template for sideloading,
 * or undefined if sideloading is not supported.
 * @param app Specifies the Office app.
 * @param addInType Specifies the type of add-in.
 */
export function getTemplatePath(
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

/**
 * Returns the web extension path in the sideload document.
 * @param app Specifies the Office app.
 * @param addInType Specifies the type of add-in.
 */
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

function isSideloadingSupportedForDesktopHost(app: OfficeApp): boolean {
  if (app === OfficeApp.Excel || app === OfficeApp.Outlook && process.platform === "win32" || app === OfficeApp.PowerPoint || app === OfficeApp.Word) {
    return true;
  }
  return false;
}

function isSideloadingSupportedForWebHost(app: OfficeApp): boolean {
  if (app === OfficeApp.Excel || app === OfficeApp.PowerPoint || app === OfficeApp.Project || app === OfficeApp.Word) {
    return true;
  }
  return false;
}

async function getOutlookExePath(): Promise<string | undefined> {
  try {
    const OutlookInstallPathRegistryKey: string = `HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\OUTLOOK.EXE`;
    const key = new registry.RegistryKey(`${OutlookInstallPathRegistryKey}`);
    const outlookExePath: string | undefined = await registry.getStringValue(key, "");

    return outlookExePath;
  } catch (err) {
    const errorMessage: string = `Unable to find Outlook install location: \n${err}`;
    throw new Error(errorMessage);
  }
}

/**
 * Given a file path, returns a unique file path where the file doesn't exist by
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

/**
 * Starts the Office app and loads the Office Add-in.
 * @param manifestPath Path to the manifest file for the Office Add-in.
 * @param app Office app to launch.
 * @param canPrompt
 */
export async function sideloadAddIn(manifestPath: string, app?: OfficeApp, canPrompt: boolean = false,
  appType?: AppType, document?: string): Promise<void> {

  try {
    if (appType === undefined) {
      appType = AppType.Desktop;
    }

    const manifest: ManifestInfo = await readManifestFile(manifestPath);
    const appsInManifest = getOfficeAppsForManifestHosts(manifest.hosts);
    const isTest: boolean = process.env.WEB_SIDELOAD_TEST !== undefined;

    if (app) {
      if (appsInManifest.indexOf(app) < 0) {
        throw new ExpectedError(`The Office Add-in manifest does not support ${getOfficeAppName(app)}.`);
      }
    } else {
      switch (appsInManifest.length) {
        case 0:
          throw new ExpectedError("The manifest does not support any Office apps.");
        case 1:
          app = appsInManifest[0];
          break;
        default:
          if (canPrompt) {
            app = await chooseOfficeApp(appsInManifest);
          }
          break;
      }
    }

    if (!app) {
      throw new ExpectedError("Please specify the Office app.");
    }

    let sideloadFile: string | undefined;

    switch (appType) {
      case AppType.Desktop:
        if (!isSideloadingSupportedForDesktopHost(app)) {
          throw new ExpectedError(`Sideload to the ${getOfficeAppName(app)} app is not supported.`);
        }

        await registerAddIn(manifestPath);

        // for Outlook, open Outlook.exe; for other Office apps, open the document
        sideloadFile = (app === OfficeApp.Outlook)
          ? await getOutlookExePath()
          : await generateSideloadFile(app, manifest, document);
        break;
      case AppType.Web:
        if (!document) {
          throw new ExpectedError(`For sideload to web, you need to specify a document url.`);
        }

        if (!isSideloadingSupportedForWebHost(app)) {
          throw new ExpectedError(`Sideload to the ${getOfficeAppName(app)} web app is not supported.`);
        }

        const manifestFileName: string = path.basename(manifestPath);
        sideloadFile = await generateSideloadUrl(manifestFileName, manifest, document, isTest);
        break;
      default:
        throw new ExpectedError("Sideload is not supported for the specified app type.");
    }

    if (sideloadFile) {
      if (app === OfficeApp.Outlook) {
        // put the Outlook.exe path in quotes if it contains spaces
        if (sideloadFile.indexOf(' ') >= 0){
          sideloadFile = `"${sideloadFile}"`;
        }

        startDetachedProcess(sideloadFile);
      } else {
        await open(sideloadFile, { wait: false });
      }
    }
    usageDataObject.reportSuccess("sideloadAddIn()");
  } catch (err) {
    usageDataObject.reportException("sideloadAddIn()", err);
  }
}