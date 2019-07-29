// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as assert from "assert";
import * as fs from "fs";
import * as mocha from "mocha";
import * as path from "path";
import * as uuid from "uuid";
import { isUUID } from "validator";
import * as manifestInfo from "../src/manifestInfo";
import {
  getOfficeAppForManifestHost,
  getOfficeAppName,
  getOfficeAppNames,
  getOfficeApps,
  getOfficeAppsForManifestHosts,
  OfficeApp,
  parseOfficeApp,
  parseOfficeApps,
  toOfficeApp,
} from "../src/officeApp";

const manifestOriginalFolder = `${process.cwd()}/test/manifests`;
const manifestTestFolder = `${process.cwd()}/testExecution/testManifests`;
const testManifest = `${manifestTestFolder}/Taskpane.manifest.xml`;

describe("Unit Tests", function() {
  describe("officeApp.ts", function() {
    describe("getOfficeAppName()", function() {
      it("Excel", function() {
        assert.strictEqual(getOfficeAppName(OfficeApp.Excel), "Excel");
      });
      it("OneNote", function() {
        assert.strictEqual(getOfficeAppName(OfficeApp.OneNote), "OneNote");
      });
      it("Outlook", function() {
        assert.strictEqual(getOfficeAppName(OfficeApp.Outlook), "Outlook");
      });
      it("PowerPoint", function() {
        assert.strictEqual(getOfficeAppName(OfficeApp.PowerPoint), "PowerPoint");
      });
      it("Project", function() {
        assert.strictEqual(getOfficeAppName(OfficeApp.Project), "Project");
      });
      it("Word", function() {
        assert.strictEqual(getOfficeAppName(OfficeApp.Word), "Word");
      });
    });
    describe("getOfficeAppNames()", function() {
      it("empty array", function() {
        const appNames = getOfficeAppNames([]);
        const [appName] = appNames;
        assert.strictEqual(appNames.length, 0);
      });
      it("one app", function() {
        const appNames = getOfficeAppNames([OfficeApp.Excel]);
        const [appName] = appNames;
        assert.strictEqual(appNames.length, 1);
        assert.strictEqual(appName, "Excel");
      });
      it("two apps", function() {
        const appNames = getOfficeAppNames([OfficeApp.Word, OfficeApp.PowerPoint]);
        const [firstAppName, secondAppName] = appNames;
        assert.strictEqual(appNames.length, 2);
        assert.strictEqual(firstAppName, "Word");
        assert.strictEqual(secondAppName, "PowerPoint");
      });
    });
    describe("getOfficeAppForManifestHost()", function() {
      it("Document", function() {
        assert.strictEqual(getOfficeAppForManifestHost("Document"), OfficeApp.Word);
        assert.strictEqual(getOfficeAppForManifestHost("document"), OfficeApp.Word);
        assert.strictEqual(getOfficeAppForManifestHost("DOCUMENT"), OfficeApp.Word);
      });
      it("Mailbox", function() {
        assert.strictEqual(getOfficeAppForManifestHost("Mailbox"), OfficeApp.Outlook);
        assert.strictEqual(getOfficeAppForManifestHost("mailbox"), OfficeApp.Outlook);
        assert.strictEqual(getOfficeAppForManifestHost("MAILBOX"), OfficeApp.Outlook);
      });
      it("Notebook", function() {
        assert.strictEqual(getOfficeAppForManifestHost("Notebook"), OfficeApp.OneNote);
        assert.strictEqual(getOfficeAppForManifestHost("notebook"), OfficeApp.OneNote);
        assert.strictEqual(getOfficeAppForManifestHost("NOTEBOOK"), OfficeApp.OneNote);
      });
      it("Presentation", function() {
        assert.strictEqual(getOfficeAppForManifestHost("Presentation"), OfficeApp.PowerPoint);
        assert.strictEqual(getOfficeAppForManifestHost("presentation"), OfficeApp.PowerPoint);
        assert.strictEqual(getOfficeAppForManifestHost("PRESENTATION"), OfficeApp.PowerPoint);
      });
      it("Project", function() {
        assert.strictEqual(getOfficeAppForManifestHost("Project"), OfficeApp.Project);
        assert.strictEqual(getOfficeAppForManifestHost("project"), OfficeApp.Project);
        assert.strictEqual(getOfficeAppForManifestHost("PROJECT"), OfficeApp.Project);
      });
      it("Workbook", function() {
        assert.strictEqual(getOfficeAppForManifestHost("Workbook"), OfficeApp.Excel);
        assert.strictEqual(getOfficeAppForManifestHost("workbook"), OfficeApp.Excel);
        assert.strictEqual(getOfficeAppForManifestHost("WORKBOOK"), OfficeApp.Excel);
      });
      it("undefined", function() {
        assert.strictEqual(getOfficeAppForManifestHost(""), undefined);
        assert.strictEqual(getOfficeAppForManifestHost("Unknown"), undefined);
      });
    });
    describe("getOfficeAppsForManifestHosts()", function() {
      it("empty array", function() {
        const apps = getOfficeAppsForManifestHosts([]);
        assert.strictEqual(apps.length, 0);
      });
      it("one host", function() {
        const apps = getOfficeAppsForManifestHosts(["Workbook"]);
        const [firstApp] = apps;
        assert.strictEqual(apps.length, 1);
        assert.strictEqual(firstApp, OfficeApp.Excel);
      });
      it("two hosts", function() {
        const apps = getOfficeAppsForManifestHosts(["Notebook", "presentation"]);
        const [firstApp, secondApp] = apps;
        assert.strictEqual(apps.length, 2);
        assert.strictEqual(firstApp, OfficeApp.OneNote);
        assert.strictEqual(secondApp, OfficeApp.PowerPoint);
      });
      it("unknown host", function() {
        const apps = getOfficeAppsForManifestHosts(["unknown"]);
        const [firstApp] = apps;
        assert.strictEqual(apps.length, 0);
      });
      it("known and unknown host", function() {
        const apps = getOfficeAppsForManifestHosts(["MailBox", "unknown"]);
        const [firstApp] = apps;
        assert.strictEqual(apps.length, 1);
        assert.strictEqual(firstApp, OfficeApp.Outlook);
      });
    });
    describe("parseOfficeApp()", function() {
      it("Excel", function() {
        assert.strictEqual(parseOfficeApp("Excel"), OfficeApp.Excel);
        assert.strictEqual(parseOfficeApp("excel"), OfficeApp.Excel);
        assert.strictEqual(parseOfficeApp("EXCEL"), OfficeApp.Excel);
      });
      it("OneNote", function() {
        assert.strictEqual(parseOfficeApp("OneNote"), OfficeApp.OneNote);
        assert.strictEqual(parseOfficeApp("onenote"), OfficeApp.OneNote);
        assert.strictEqual(parseOfficeApp("ONENOTE"), OfficeApp.OneNote);
      });
      it("Outlook", function() {
        assert.strictEqual(parseOfficeApp("Outlook"), OfficeApp.Outlook);
        assert.strictEqual(parseOfficeApp("outlook"), OfficeApp.Outlook);
        assert.strictEqual(parseOfficeApp("OUTLOOK"), OfficeApp.Outlook);
      });
      it("PowerPoint", function() {
        assert.strictEqual(parseOfficeApp("PowerPoint"), OfficeApp.PowerPoint);
        assert.strictEqual(parseOfficeApp("powerpoint"), OfficeApp.PowerPoint);
        assert.strictEqual(parseOfficeApp("POWERPOINT"), OfficeApp.PowerPoint);
      });
      it("Project", function() {
        assert.strictEqual(parseOfficeApp("Project"), OfficeApp.Project);
        assert.strictEqual(parseOfficeApp("project"), OfficeApp.Project);
        assert.strictEqual(parseOfficeApp("PROJECT"), OfficeApp.Project);
      });
      it("Word", function() {
        assert.strictEqual(parseOfficeApp("Word"), OfficeApp.Word);
        assert.strictEqual(parseOfficeApp("word"), OfficeApp.Word);
        assert.strictEqual(parseOfficeApp("WORD"), OfficeApp.Word);
      });
      it("should trim whitespace", function() {
        assert.strictEqual(parseOfficeApp(" excel"), OfficeApp.Excel);
        assert.strictEqual(parseOfficeApp("word\n"), OfficeApp.Word);
        assert.strictEqual(parseOfficeApp("  \toutlook  "), OfficeApp.Outlook);
      });
    });
    describe("parseOfficeApps()", function() {
      it("one app", function() {
        const apps = parseOfficeApps("excel");
        const [firstApp] = apps;
        assert.strictEqual(apps.length, 1);
        assert.strictEqual(firstApp, OfficeApp.Excel);
      });
      it("two apps", function() {
        const apps = parseOfficeApps("OneNote,PowerPoint");
        const [firstApp, secondApp] = apps;
        assert.strictEqual(apps.length, 2);
        assert.strictEqual(firstApp, OfficeApp.OneNote);
        assert.strictEqual(secondApp, OfficeApp.PowerPoint);
      });
      it("two apps with whitespace", function() {
        const apps = parseOfficeApps(" OneNote, PowerPoint ");
        const [firstApp, secondApp] = apps;
        assert.strictEqual(apps.length, 2);
        assert.strictEqual(firstApp, OfficeApp.OneNote);
        assert.strictEqual(secondApp, OfficeApp.PowerPoint);
      });
      it("all", function() {
        const apps = parseOfficeApps("all");
        assert.strictEqual(apps.length, 6);
        assert.strictEqual(apps[0], OfficeApp.Excel);
        assert.strictEqual(apps[1], OfficeApp.OneNote);
        assert.strictEqual(apps[2], OfficeApp.Outlook);
        assert.strictEqual(apps[3], OfficeApp.PowerPoint);
        assert.strictEqual(apps[4], OfficeApp.Project);
        assert.strictEqual(apps[5], OfficeApp.Word);
        const appsWithWhitespace = parseOfficeApps(" all ");
        assert.strictEqual(appsWithWhitespace.length, 6);
        assert.strictEqual(appsWithWhitespace[0], OfficeApp.Excel);
        assert.strictEqual(appsWithWhitespace[1], OfficeApp.OneNote);
        assert.strictEqual(appsWithWhitespace[2], OfficeApp.Outlook);
        assert.strictEqual(appsWithWhitespace[3], OfficeApp.PowerPoint);
        assert.strictEqual(appsWithWhitespace[4], OfficeApp.Project);
        assert.strictEqual(appsWithWhitespace[5], OfficeApp.Word);
      });
      it("unknown app", function() {
        const unknown = "unknown";
        assert.throws(() => {
          parseOfficeApps(unknown);
        }, new Error(`"${unknown}" is not a valid Office app.`));
        assert.throws(() => {
          parseOfficeApps(`Excel,${unknown}`);
        }, new Error(`"${unknown}" is not a valid Office app.`));
      });
      it("empty string", function() {
        assert.throws(() => {
          parseOfficeApps("");
        }, new Error(`"" is not a valid Office app.`));
      });
    });
    describe("toOfficeApp()", function() {
      it("Excel", function() {
        assert.strictEqual(toOfficeApp("Excel"), OfficeApp.Excel);
        assert.strictEqual(toOfficeApp("excel"), OfficeApp.Excel);
        assert.strictEqual(toOfficeApp("EXCEL"), OfficeApp.Excel);
      });
      it("OneNote", function() {
        assert.strictEqual(toOfficeApp("OneNote"), OfficeApp.OneNote);
        assert.strictEqual(toOfficeApp("onenote"), OfficeApp.OneNote);
        assert.strictEqual(toOfficeApp("ONENOTE"), OfficeApp.OneNote);
      });
      it("Outlook", function() {
        assert.strictEqual(toOfficeApp("Outlook"), OfficeApp.Outlook);
        assert.strictEqual(toOfficeApp("outlook"), OfficeApp.Outlook);
        assert.strictEqual(toOfficeApp("OUTLOOK"), OfficeApp.Outlook);
      });
      it("PowerPoint", function() {
        assert.strictEqual(toOfficeApp("PowerPoint"), OfficeApp.PowerPoint);
        assert.strictEqual(toOfficeApp("powerpoint"), OfficeApp.PowerPoint);
        assert.strictEqual(toOfficeApp("POWERPOINT"), OfficeApp.PowerPoint);
      });
      it("Project", function() {
        assert.strictEqual(toOfficeApp("Project"), OfficeApp.Project);
        assert.strictEqual(toOfficeApp("project"), OfficeApp.Project);
        assert.strictEqual(toOfficeApp("PROJECT"), OfficeApp.Project);
      });
      it("Word", function() {
        assert.strictEqual(toOfficeApp("Word"), OfficeApp.Word);
        assert.strictEqual(toOfficeApp("word"), OfficeApp.Word);
        assert.strictEqual(toOfficeApp("WORD"), OfficeApp.Word);
      });
      it("should return undefined for an unknown or empty value", function() {
        assert.strictEqual(toOfficeApp("unknown"), undefined);
        assert.strictEqual(toOfficeApp(""), undefined);
      });
      it("should trim whitespace", function() {
        assert.strictEqual(toOfficeApp(" excel"), OfficeApp.Excel);
        assert.strictEqual(toOfficeApp("word\n"), OfficeApp.Word);
        assert.strictEqual(toOfficeApp("  \toutlook  "), OfficeApp.Outlook);
        assert.strictEqual(toOfficeApp("  unknown  "), undefined);
        assert.strictEqual(toOfficeApp("    "), undefined);
      });
    });
    describe("getOfficeApps()", function() {
      it("should return all Office apps", function() {
        const apps = getOfficeApps();
        assert.strictEqual(apps.length, 6);
        assert.strictEqual(apps[0], OfficeApp.Excel);
        assert.strictEqual(apps[1], OfficeApp.OneNote);
        assert.strictEqual(apps[2], OfficeApp.Outlook);
        assert.strictEqual(apps[3], OfficeApp.PowerPoint);
        assert.strictEqual(apps[4], OfficeApp.Project);
        assert.strictEqual(apps[5], OfficeApp.Word);
      });
    });
  });
  describe("manifestInfo.ts", function() {
    describe("readManifestInfo()", function() {
      it("should read the manifest info", async function() {
        const info = await manifestInfo.readManifestFile("test/manifests/TaskPane.manifest.xml");

        assert.strictEqual(info.allowSnapshot, undefined);
        assert.strictEqual(info.alternateId, undefined);
        assert.strictEqual(info.appDomains instanceof Array, true);
        assert.strictEqual(info.appDomains!.length, 1);
        assert.strictEqual(info.appDomains![0], "contoso.com");
        assert.strictEqual(info.defaultLocale, "en-US");
        assert.strictEqual(info.description, "Describes this Office Add-in.");
        assert.strictEqual(info.displayName, "Contoso Task Pane Add-in");
        assert.strictEqual(info.highResolutionIconUrl, "https://localhost:3000/assets/icon-80.png");
        assert.strictEqual(info.hosts instanceof Array, true);
        assert.strictEqual(info.hosts!.length, 1);
        assert.strictEqual(info.hosts![0], "Workbook");
        assert.strictEqual(info.iconUrl, "https://localhost:3000/assets/icon-32.png");
        assert.strictEqual(info.id, "6c883c79-9b2a-45a3-b3d1-3dbd08000c5a");
        assert.strictEqual(info.officeAppType, "TaskPaneApp");
        assert.strictEqual(info.permissions, "ReadWriteDocument");
        assert.strictEqual(info.providerName, "Contoso");
        assert.strictEqual(info.supportUrl, "https://www.contoso.com/help");
        assert.strictEqual(info.version, "1.2.3.4");

        if (info.defaultSettings) {
          assert.strictEqual(info.defaultSettings.sourceLocation, "https://localhost:3000/taskpane.html");
        } else {
          assert.fail("ManifestInfo should include defaultSettings.");
        }
      });
      it("should throw an error if there is a bad xml end tag", async function() {
          let result;
          try {
            await manifestInfo.readManifestFile("test/manifests/manifest.incorrect-end-tag.xml");
          } catch (err) {
            result = err;
          }
          assert.equal(result.message, "Unable to parse the manifest file: test/manifests/manifest.incorrect-end-tag.xml. \nError: Unexpected close tag\nLine: 8\nColumn: 46\nChar: >");
      });
      it ("should handle OfficeApp with no info", async function() {
        const info = await manifestInfo.readManifestFile("test/manifests/manifest.officeapp-empty.xml");

        assert.strictEqual(info.allowSnapshot, undefined);
        assert.strictEqual(info.alternateId, undefined);
        assert.strictEqual(info.appDomains, undefined);
        assert.strictEqual(info.defaultLocale, undefined);
        assert.strictEqual(info.defaultSettings, undefined);
        assert.strictEqual(info.description, undefined);
        assert.strictEqual(info.displayName, undefined);
        assert.strictEqual(info.highResolutionIconUrl, undefined);
        assert.strictEqual(info.hosts, undefined);
        assert.strictEqual(info.iconUrl, undefined);
        assert.strictEqual(info.id, undefined);
        assert.strictEqual(info.officeAppType, undefined);
        assert.strictEqual(info.permissions, undefined);
        assert.strictEqual(info.providerName, undefined);
        assert.strictEqual(info.supportUrl, undefined);
        assert.strictEqual(info.version, undefined);
      });
      it("should handle a missing description", async function() {
        const info = await manifestInfo.readManifestFile("test/manifests/manifest.no-description.xml");

        assert.strictEqual(info.defaultLocale, "en-US");
        assert.strictEqual(info.description, undefined);
        assert.strictEqual(info.displayName, "Office Add-in Name");
        assert.strictEqual(info.id, "132a8a21-011a-4ceb-9336-6af8a276a288");
        assert.strictEqual(info.officeAppType, "TaskPaneApp");
        assert.strictEqual(info.providerName, "ProviderName");
        assert.strictEqual(info.version, "1.2.3.4");
      });
    });
    describe("modifyManifestFile()", function() {
      beforeEach(async function() {
        await _createManifestFilesFolder();
      });
      afterEach(async function() {
        await _deleteManifestTestFolder(manifestTestFolder);
      });
      it("should handle a specified valid guid and displayName", async function() {
        // call modify, specifying guid and displayName  parameters
        const testGuid = uuid.v1();
        const testDisplayName = "TestDisplayName";
        const updatedInfo = await manifestInfo.modifyManifestFile(testManifest, testGuid, testDisplayName);

        // verify guid displayName updated
        assert.strictEqual(updatedInfo.id, testGuid);
        assert.strictEqual(updatedInfo.displayName, testDisplayName);
      });
      it(`should handle specifying "random" form guid parameter`, async function() {
        // get original manifest info and create copy of manifest that we can overwrite in this test
        const originalInfo = await manifestInfo.readManifestFile(testManifest);

        // call modify, specifying "random" parameter
        const updatedInfo = await manifestInfo.modifyManifestFile(testManifest, "random", undefined);

        // verify guid updated, that it"s a valid guid and that the displayName is not updated
        assert.notStrictEqual(updatedInfo.id, originalInfo.id);
        assert.strictEqual(updatedInfo.id && isUUID(updatedInfo.id), true);
        assert.strictEqual(updatedInfo.displayName, originalInfo.displayName);
      });
      it("should handle specifying displayName only", async function() {
        // get original manifest info and create copy of manifest that we can overwrite in this test
        const originalInfo = await manifestInfo.readManifestFile(testManifest);

        // call  modify, specifying a displayName parameter
        const testDisplayName = "TestDisplayName";
        const updatedInfo = await manifestInfo.modifyManifestFile(testManifest, undefined, testDisplayName);

        // verify displayName updated and guid not updated
        assert.notStrictEqual(updatedInfo.displayName, originalInfo.displayName);
        assert.strictEqual(updatedInfo.displayName, testDisplayName);
        assert.strictEqual( updatedInfo.id, originalInfo.id);
      });
      it("should handle not specifying either a guid or displayName", async function() {
        let result;
        try {
          await manifestInfo.modifyManifestFile(testManifest);
        } catch (err) {
          result = err.message;
        }
        assert.strictEqual(result, `You need to specify something to change in the manifest.`);
      });
      it("should handle an invalid manifest file path", async function() {
        // call  modify, specifying an invalid manifest path with a valid guid and displayName
        const invalidManifest = path.normalize(`${manifestTestFolder}/foo/manifest.xml`);
        const testGuid = uuid.v1();
        const testDisplayName = "TestDisplayName";
        let result;
        try {
          await manifestInfo.modifyManifestFile(invalidManifest, testGuid, testDisplayName);
        } catch (err) {
          result = err.message;
        }

        assert.strictEqual(result, `Unable to modify xml data for manifest file: ${invalidManifest}. \nError: ENOENT: no such file or directory, open '${invalidManifest}'`);
      });
    });
  });
});

async function _deleteManifestTestFolder(projectFolder: string): Promise<void> {
  if (fs.existsSync(projectFolder)) {
    fs.readdirSync(projectFolder).forEach(function(file) {
    const curPath = projectFolder + "/" + file;

    if (fs.lstatSync(curPath).isDirectory()) {
      _deleteManifestTestFolder(curPath);
    } else {
      fs.unlinkSync(curPath);
    }
  });
    fs.rmdirSync(projectFolder);
  }
}

async function _createManifestFilesFolder(): Promise<void> {
    if (fs.existsSync(manifestTestFolder)) {
      _deleteManifestTestFolder(manifestTestFolder);
  }
    const fsExtra = require("fs-extra");
    await fsExtra.copy(manifestOriginalFolder, manifestTestFolder);
}
