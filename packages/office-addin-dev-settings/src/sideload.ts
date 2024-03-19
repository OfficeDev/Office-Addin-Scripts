// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as fs from "fs";
import * as childProcess from "child_process";
import * as AdmZip from "adm-zip";
import {
  AddInType,
  getAddInTypeForManifestOfficeAppType,
  getOfficeAppName,
  getOfficeAppsForManifestHosts,
  ManifestInfo,
  OfficeApp,
  OfficeAddinManifest,
} from "office-addin-manifest";
import open = require("open");
import semver = require("semver");
import * as os from "os";
import * as path from "path";
import { AppType } from "./appType";
import { registerAddIn } from "./dev-settings";
import { startDetachedProcess } from "./process";
import { chooseOfficeApp } from "./prompt";
import * as registry from "./registry";
import { usageDataObject } from "./defaults";
import { ExpectedError } from "office-addin-usage-data";

/* global __dirname, Buffer, console, process, URL */

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
    throw new ExpectedError("The manifest does not contain the version for the add-in.");
  }
  const addInType = getAddInTypeForManifestOfficeAppType(manifest.officeAppType);

  if (!addInType) {
    throw new ExpectedError("The manifest contains an unsupported OfficeApp xsi:type.");
  }

  const documentWasProvided = document && document !== "";
  const templatePath = documentWasProvided ? path.resolve(document) : getTemplatePath(app, addInType);

  if (!templatePath) {
    throw new ExpectedError(`Sideload is not supported for apptype: ${addInType}.`);
  }

  const appName = getOfficeAppName(app);
  const extension = path.extname(templatePath);
  const pathToWrite: string = makePathUnique(
    path.join(os.tmpdir(), `${appName} add-in ${manifest.id}${extension}`),
    true
  );

  if (!documentWasProvided) {
    const webExtensionPath = getWebExtensionPath(app, addInType);
    if (!webExtensionPath) {
      throw new ExpectedError("Don't know the webextension path.");
    }

    // replace the placeholder id and version
    const templateZip: AdmZip = new AdmZip(templatePath);
    const outZip: AdmZip = new AdmZip();
    const extEntry = templateZip.getEntry(webExtensionPath);

    if (!extEntry) {
      throw new ExpectedError("webextension was not found.");
    }

    const webExtensionXml = templateZip
      .readAsText(extEntry)
      .replace(/00000000-0000-0000-0000-000000000000/g, manifest.id)
      .replace(/1.0.0.0/g, manifest.version);

    templateZip.getEntries().forEach(function (entry) {
      var data: Buffer = entry.getData();
      if (entry == extEntry) {
        data = Buffer.from(webExtensionXml);
      }
      outZip.addFile(entry.entryName, data, entry.comment, entry.attr);
    });

    // Write the file
    await outZip.writeZipPromise(pathToWrite);
  } else {
    await fs.promises.copyFile(templatePath, pathToWrite);
  }

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
export async function generateSideloadUrl(
  manifestFileName: string,
  manifest: ManifestInfo,
  documentUrl: string,
  isTest: boolean = false
): Promise<string> {
  const testQueryParam = "&wdaddintest=true";

  if (!manifest.id) {
    throw new ExpectedError("The manifest does not contain the id for the add-in.");
  }

  if (manifest.defaultSettings === undefined || manifest.defaultSettings.sourceLocation === undefined) {
    throw new ExpectedError("The manifest does not contain the SourceLocation for the add-in");
  }

  const sourceLocationUrl: URL = new URL(manifest.defaultSettings.sourceLocation);
  if (sourceLocationUrl.protocol.indexOf("https") === -1) {
    throw new ExpectedError("The SourceLocation in the manifest does not use the HTTPS protocol.");
  }

  if (sourceLocationUrl.host.indexOf("localhost") === -1 && sourceLocationUrl.host.indexOf("127.0.0.1") === -1) {
    throw new ExpectedError(
      "The hostname specified by the SourceLocation in the manifest is not supported for sideload. The hostname should be 'localhost' or 127.0.0.1."
    );
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
export function getTemplatePath(app: OfficeApp, addInType: AddInType): string | undefined {
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
function getWebExtensionPath(app: OfficeApp, addInType: AddInType): string | undefined {
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
  if (
    app === OfficeApp.Excel ||
    (app === OfficeApp.Outlook && process.platform === "win32") ||
    app === OfficeApp.PowerPoint ||
    app === OfficeApp.Word
  ) {
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

function hasOfficeVersion(targetVersion: string, currentVersion: string): boolean {
  return semver.gte(currentVersion, targetVersion);
}

async function getOutlookVersion(): Promise<string | undefined> {
  try {
    const key = new registry.RegistryKey(`HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Office\\ClickToRun\\Configuration`);
    const outlookInstallVersion: string | undefined = await registry.getStringValue(key, "ClientVersionToReport");
    const outlookSmallerVersion = outlookInstallVersion?.split(`.`, 3).join(`.`);

    return outlookSmallerVersion;
  } catch (err) {
    return undefined;
  }
}

async function getOfficeExePath(app: OfficeApp): Promise<string> {
  let hostApp: string = "";
  try {
    switch (app) {
      case OfficeApp.Excel:
        hostApp = "excel.exe";
        break;
      case OfficeApp.Outlook:
        hostApp = "OUTLOOK.EXE";
        break;
      case OfficeApp.Word:
        hostApp = "Winword.exe";
        break;
      case OfficeApp.PowerPoint:
        hostApp = "powerpnt.exe";
        break;
      default:
        hostApp = "OUTLOOK.EXE";
        break;
    }

    const InstallPathRegistryKey: string = `HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\${hostApp}`;
    const key = new registry.RegistryKey(`${InstallPathRegistryKey}`);
    const ExePath: string | undefined = await registry.getStringValue(key, "");

    if (!ExePath) {
      throw new Error(`${hostApp} registry empty`);
    }
    return ExePath;
  } catch (err) {
    const errorMessage: string = `Unable to find "${hostApp}" install location: \n${err}`;    
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

      currentPath = path.join(parsedPath.dir, `${parsedPath.name}.${suffix}${parsedPath.ext}`);
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
export async function sideloadAddIn(
  manifestPath: string,
  app?: OfficeApp,
  canPrompt: boolean = false,
  appType?: AppType,
  document?: string,
  registration?: string
): Promise<void> {
  try {
    if (appType === undefined) {
      appType = AppType.Desktop;
    }

    const manifest: ManifestInfo = await OfficeAddinManifest.readManifestFile(manifestPath);
    const appsInManifest: OfficeApp[] = getOfficeAppsForManifestHosts(manifest.hosts);

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

    switch (appType) {
      case AppType.Desktop:
        await registerAddIn(manifestPath, registration);
        await launchDesktopApp(app, manifestPath, manifest, document);
        break;
      case AppType.Web: {
        if (!document) {
          throw new ExpectedError(`For sideload to web, you need to specify a document url.`);
        }
        await launchWebApp(app, manifestPath, manifest, document);
        break;
      }
      default:
        throw new ExpectedError("Sideload is not supported for the specified app type.");
    }
    usageDataObject.reportSuccess("sideloadAddIn()");
  } catch (err: any) {
    usageDataObject.reportException("sideloadAddIn()", err);
    throw err;
  }
}

async function launchDesktopApp(app: OfficeApp, manifestPath: string, manifest: ManifestInfo, document?: string) {
  if (!isSideloadingSupportedForDesktopHost(app)) {
    throw new ExpectedError(`Sideload to the ${getOfficeAppName(app)} app is not supported.`);
  }

  // for Outlook, Word, Excel, PowerPoint open {Host}.exe; for other Office apps, open the document
  let path: string;
  if (app == OfficeApp.Outlook) {
    const version: string | undefined = await getOutlookVersion();
    if (version && !hasOfficeVersion("16.0.13709", version)) {
      throw new ExpectedError(
        `The current version of Outlook does not support sideload. Please use version 16.0.13709 or greater.`
      );
    }
    path = await getOfficeExePath(app);
  } else if (manifestPath.endsWith(".json")) {
    path = await getOfficeExePath(app);
  } else {
    path = await generateSideloadFile(app, manifest, document);
  }

  await launchApp(app, path);
}

async function launchWebApp(app: OfficeApp, manifestPath: string, manifest: ManifestInfo, document: string) {
  if (!isSideloadingSupportedForWebHost(app)) {
    throw new ExpectedError(`Sideload to the ${getOfficeAppName(app)} web app is not supported.`);
  }
  const manifestFileName: string = path.basename(manifestPath);
  const isTest: boolean = process.env.WEB_SIDELOAD_TEST !== undefined;
  await launchApp(app, await generateSideloadUrl(manifestFileName, manifest, document, isTest));
}

async function launchApp(app: OfficeApp, sideloadFile: string) {
  console.log(`Launching ${app} via ${sideloadFile}`);
  if (sideloadFile) {
    if (app === OfficeApp.Outlook) {
      // put the Outlook.exe path in quotes if it contains spaces
      if (sideloadFile.indexOf(" ") >= 0) {
        sideloadFile = `"${sideloadFile}"`;
      }

      startDetachedProcess(sideloadFile);
    } else {
      await open(sideloadFile, { wait: false });
    }
  }
}
