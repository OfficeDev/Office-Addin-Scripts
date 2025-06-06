// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import AdmZip from "adm-zip";
import assert from "assert";
import fs from "fs";
import { afterEach, beforeEach, describe, it } from "mocha";
import path from "path";
import { v1 as uuidv1 } from "uuid";
import { isUUID } from "validator";
import {
  AddInType,
  getAddInTypeForManifestOfficeAppType,
  getAddInTypes,
  parseAddInType,
  parseAddInTypes,
  toAddInType,
} from "../src/addInTypes";
import { OfficeAddinManifest } from "../src/manifestOperations";
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
import { validateManifest } from "../src/validate";
import { exportMetadataPackage } from "../src/export";
import { ManifestInfo } from "../src/manifestInfo";

/* global require */

const manifestOriginalFolder = path.resolve("./test/manifests");
const manifestTestFolder = path.resolve("./testExecution/testManifests");
const testManifestXml = path.resolve(manifestTestFolder, "TaskPane.manifest.xml");
const testManifestJson = path.resolve(manifestTestFolder, "manifest.json");

describe("Unit Tests", function () {
  describe("addInTypes.ts", function () {
    describe("getAddInTypeForManifestOfficeAppType()", function () {
      it("Content", function () {
        assert.strictEqual(getAddInTypeForManifestOfficeAppType("ContentApp"), AddInType.Content);
        assert.strictEqual(getAddInTypeForManifestOfficeAppType("contentapp"), AddInType.Content);
        assert.strictEqual(getAddInTypeForManifestOfficeAppType(" contentApp "), AddInType.Content);
      });
      it("Mail", function () {
        assert.strictEqual(getAddInTypeForManifestOfficeAppType("MailApp"), AddInType.Mail);
        assert.strictEqual(getAddInTypeForManifestOfficeAppType("mailapp"), AddInType.Mail);
        assert.strictEqual(getAddInTypeForManifestOfficeAppType(" mailApp "), AddInType.Mail);
      });
      it("TaskPane", function () {
        assert.strictEqual(getAddInTypeForManifestOfficeAppType("TaskPaneApp"), AddInType.TaskPane);
        assert.strictEqual(getAddInTypeForManifestOfficeAppType("taskpaneapp"), AddInType.TaskPane);
        assert.strictEqual(
          getAddInTypeForManifestOfficeAppType(" taskpaneApp "),
          AddInType.TaskPane
        );
      });
      it("unknown", function () {
        assert.strictEqual(getAddInTypeForManifestOfficeAppType("Unknown"), undefined);
        assert.strictEqual(getAddInTypeForManifestOfficeAppType(""), undefined);
      });
    });
    describe("getAddInTypes()", function () {
      it("should return all add-in types", function () {
        const types = getAddInTypes();
        assert.strictEqual(types.length, 3);
        assert.strictEqual(types[0], AddInType.Content);
        assert.strictEqual(types[1], AddInType.Mail);
        assert.strictEqual(types[2], AddInType.TaskPane);
      });
    });
    describe("parseAddInType()", function () {
      it("Content", function () {
        assert.strictEqual(parseAddInType("content"), AddInType.Content);
        assert.strictEqual(parseAddInType("Content"), AddInType.Content);
        assert.strictEqual(parseAddInType(" CONTENT "), AddInType.Content);
      });
      it("Mail", function () {
        assert.strictEqual(parseAddInType("mail"), AddInType.Mail);
        assert.strictEqual(parseAddInType("Mail"), AddInType.Mail);
        assert.strictEqual(parseAddInType(" MAIL "), AddInType.Mail);
      });
      it("TaskPane", function () {
        assert.strictEqual(parseAddInType("taskpane"), AddInType.TaskPane);
        assert.strictEqual(parseAddInType("TaskPane"), AddInType.TaskPane);
        assert.strictEqual(parseAddInType(" TASKPANE "), AddInType.TaskPane);
      });
    });
    describe("parseAddInTypes()", function () {
      it("one type", function () {
        const types = parseAddInTypes("taskpane");
        const [firstType] = types;
        assert.strictEqual(types.length, 1);
        assert.strictEqual(firstType, AddInType.TaskPane);
      });
      it("two types", function () {
        const types = parseAddInTypes("Mail,Content");
        const [first, second] = types;
        assert.strictEqual(types.length, 2);
        assert.strictEqual(first, AddInType.Mail);
        assert.strictEqual(second, AddInType.Content);
      });
      it("two types with whitespace", function () {
        const types = parseAddInTypes(" TaskPane, Content ");
        const [first, second] = types;
        assert.strictEqual(types.length, 2);
        assert.strictEqual(first, AddInType.TaskPane);
        assert.strictEqual(second, AddInType.Content);
      });
      it("all", function () {
        const types = parseAddInTypes("all");
        assert.strictEqual(types.length, 3);
        assert.strictEqual(types[0], AddInType.Content);
        assert.strictEqual(types[1], AddInType.Mail);
        assert.strictEqual(types[2], AddInType.TaskPane);
        const typesWithWhitespace = parseAddInTypes(" all ");
        assert.strictEqual(typesWithWhitespace.length, 3);
        assert.strictEqual(typesWithWhitespace[0], AddInType.Content);
        assert.strictEqual(typesWithWhitespace[1], AddInType.Mail);
        assert.strictEqual(typesWithWhitespace[2], AddInType.TaskPane);
      });
      it("unknown app", function () {
        const unknown = "unknown";
        assert.throws(
          () => {
            parseOfficeApps(unknown);
          },
          new Error(`"${unknown}" is not a valid Office app.`)
        );
        assert.throws(
          () => {
            parseOfficeApps(`Excel,${unknown}`);
          },
          new Error(`"${unknown}" is not a valid Office app.`)
        );
      });
      it("empty string", function () {
        assert.throws(() => {
          parseOfficeApps("");
        }, new Error(`"" is not a valid Office app.`));
      });
    });
    describe("toAddInType()", function () {
      it("Content", function () {
        assert.strictEqual(toAddInType("Content"), AddInType.Content);
        assert.strictEqual(toAddInType("content"), AddInType.Content);
      });
      it("Mail", function () {
        assert.strictEqual(toAddInType("Mail"), AddInType.Mail);
        assert.strictEqual(toAddInType("mail"), AddInType.Mail);
      });
      it("TaskPane", function () {
        assert.strictEqual(toAddInType("TaskPane"), AddInType.TaskPane);
        assert.strictEqual(toAddInType("taskpane"), AddInType.TaskPane);
      });
      it("should return undefined for an unknown or empty value", function () {
        assert.strictEqual(toAddInType("unknown"), undefined);
        assert.strictEqual(toAddInType(""), undefined);
      });
      it("should trim whitespace", function () {
        assert.strictEqual(toAddInType(" taskPane "), AddInType.TaskPane);
      });
    });
  });
  describe("officeApp.ts", function () {
    describe("getOfficeAppForManifestHost()", function () {
      it("Document", function () {
        assert.strictEqual(getOfficeAppForManifestHost("Document"), OfficeApp.Word);
        assert.strictEqual(getOfficeAppForManifestHost("document"), OfficeApp.Word);
        assert.strictEqual(getOfficeAppForManifestHost("DOCUMENT"), OfficeApp.Word);
      });
      it("Mail", function () {
        assert.strictEqual(getOfficeAppForManifestHost("Mail"), OfficeApp.Outlook);
        assert.strictEqual(getOfficeAppForManifestHost("mail"), OfficeApp.Outlook);
        assert.strictEqual(getOfficeAppForManifestHost("MAIL"), OfficeApp.Outlook);
      });
      it("Mailbox", function () {
        assert.strictEqual(getOfficeAppForManifestHost("Mailbox"), OfficeApp.Outlook);
        assert.strictEqual(getOfficeAppForManifestHost("mailbox"), OfficeApp.Outlook);
        assert.strictEqual(getOfficeAppForManifestHost("MAILBOX"), OfficeApp.Outlook);
      });
      it("Notebook", function () {
        assert.strictEqual(getOfficeAppForManifestHost("Notebook"), OfficeApp.OneNote);
        assert.strictEqual(getOfficeAppForManifestHost("notebook"), OfficeApp.OneNote);
        assert.strictEqual(getOfficeAppForManifestHost("NOTEBOOK"), OfficeApp.OneNote);
      });
      it("Presentation", function () {
        assert.strictEqual(getOfficeAppForManifestHost("Presentation"), OfficeApp.PowerPoint);
        assert.strictEqual(getOfficeAppForManifestHost("presentation"), OfficeApp.PowerPoint);
        assert.strictEqual(getOfficeAppForManifestHost("PRESENTATION"), OfficeApp.PowerPoint);
      });
      it("Project", function () {
        assert.strictEqual(getOfficeAppForManifestHost("Project"), OfficeApp.Project);
        assert.strictEqual(getOfficeAppForManifestHost("project"), OfficeApp.Project);
        assert.strictEqual(getOfficeAppForManifestHost("PROJECT"), OfficeApp.Project);
      });
      it("Workbook", function () {
        assert.strictEqual(getOfficeAppForManifestHost("Workbook"), OfficeApp.Excel);
        assert.strictEqual(getOfficeAppForManifestHost("workbook"), OfficeApp.Excel);
        assert.strictEqual(getOfficeAppForManifestHost("WORKBOOK"), OfficeApp.Excel);
      });
      it("undefined", function () {
        assert.strictEqual(getOfficeAppForManifestHost(""), undefined);
        assert.strictEqual(getOfficeAppForManifestHost("Unknown"), undefined);
      });
    });
    describe("getOfficeAppName()", function () {
      it("Excel", function () {
        assert.strictEqual(getOfficeAppName(OfficeApp.Excel), "Excel");
      });
      it("OneNote", function () {
        assert.strictEqual(getOfficeAppName(OfficeApp.OneNote), "OneNote");
      });
      it("Outlook", function () {
        assert.strictEqual(getOfficeAppName(OfficeApp.Outlook), "Outlook");
      });
      it("PowerPoint", function () {
        assert.strictEqual(getOfficeAppName(OfficeApp.PowerPoint), "PowerPoint");
      });
      it("Project", function () {
        assert.strictEqual(getOfficeAppName(OfficeApp.Project), "Project");
      });
      it("Word", function () {
        assert.strictEqual(getOfficeAppName(OfficeApp.Word), "Word");
      });
    });
    describe("getOfficeAppNames()", function () {
      it("empty array", function () {
        const appNames = getOfficeAppNames([]);
        assert.strictEqual(appNames.length, 0);
      });
      it("one app", function () {
        const appNames = getOfficeAppNames([OfficeApp.Excel]);
        const [appName] = appNames;
        assert.strictEqual(appNames.length, 1);
        assert.strictEqual(appName, "Excel");
      });
      it("two apps", function () {
        const appNames = getOfficeAppNames([OfficeApp.Word, OfficeApp.PowerPoint]);
        const [firstAppName, secondAppName] = appNames;
        assert.strictEqual(appNames.length, 2);
        assert.strictEqual(firstAppName, "Word");
        assert.strictEqual(secondAppName, "PowerPoint");
      });
    });
    describe("getOfficeApps()", function () {
      it("should return all Office apps", function () {
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
    describe("getOfficeAppsForManifestHosts()", function () {
      it("empty array", function () {
        const apps = getOfficeAppsForManifestHosts([]);
        assert.strictEqual(apps.length, 0);
      });
      it("one host", function () {
        const apps = getOfficeAppsForManifestHosts(["Workbook"]);
        const [firstApp] = apps;
        assert.strictEqual(apps.length, 1);
        assert.strictEqual(firstApp, OfficeApp.Excel);
      });
      it("two hosts", function () {
        const apps = getOfficeAppsForManifestHosts(["Notebook", "presentation"]);
        const [firstApp, secondApp] = apps;
        assert.strictEqual(apps.length, 2);
        assert.strictEqual(firstApp, OfficeApp.OneNote);
        assert.strictEqual(secondApp, OfficeApp.PowerPoint);
      });
      it("unknown host", function () {
        const apps = getOfficeAppsForManifestHosts(["unknown"]);
        assert.strictEqual(apps.length, 0);
      });
      it("known and unknown host", function () {
        const apps = getOfficeAppsForManifestHosts(["MailBox", "unknown"]);
        const [firstApp] = apps;
        assert.strictEqual(apps.length, 1);
        assert.strictEqual(firstApp, OfficeApp.Outlook);
      });
    });
    describe("parseOfficeApp()", function () {
      it("Excel", function () {
        assert.strictEqual(parseOfficeApp("Excel"), OfficeApp.Excel);
        assert.strictEqual(parseOfficeApp("excel"), OfficeApp.Excel);
        assert.strictEqual(parseOfficeApp("EXCEL"), OfficeApp.Excel);
      });
      it("OneNote", function () {
        assert.strictEqual(parseOfficeApp("OneNote"), OfficeApp.OneNote);
        assert.strictEqual(parseOfficeApp("onenote"), OfficeApp.OneNote);
        assert.strictEqual(parseOfficeApp("ONENOTE"), OfficeApp.OneNote);
      });
      it("Outlook", function () {
        assert.strictEqual(parseOfficeApp("Outlook"), OfficeApp.Outlook);
        assert.strictEqual(parseOfficeApp("outlook"), OfficeApp.Outlook);
        assert.strictEqual(parseOfficeApp("OUTLOOK"), OfficeApp.Outlook);
      });
      it("PowerPoint", function () {
        assert.strictEqual(parseOfficeApp("PowerPoint"), OfficeApp.PowerPoint);
        assert.strictEqual(parseOfficeApp("powerpoint"), OfficeApp.PowerPoint);
        assert.strictEqual(parseOfficeApp("POWERPOINT"), OfficeApp.PowerPoint);
      });
      it("Project", function () {
        assert.strictEqual(parseOfficeApp("Project"), OfficeApp.Project);
        assert.strictEqual(parseOfficeApp("project"), OfficeApp.Project);
        assert.strictEqual(parseOfficeApp("PROJECT"), OfficeApp.Project);
      });
      it("Word", function () {
        assert.strictEqual(parseOfficeApp("Word"), OfficeApp.Word);
        assert.strictEqual(parseOfficeApp("word"), OfficeApp.Word);
        assert.strictEqual(parseOfficeApp("WORD"), OfficeApp.Word);
      });
      it("should trim whitespace", function () {
        assert.strictEqual(parseOfficeApp(" excel"), OfficeApp.Excel);
        assert.strictEqual(parseOfficeApp("word\n"), OfficeApp.Word);
        assert.strictEqual(parseOfficeApp("  \toutlook  "), OfficeApp.Outlook);
      });
    });
    describe("parseOfficeApps()", function () {
      it("one app", function () {
        const apps = parseOfficeApps("excel");
        const [firstApp] = apps;
        assert.strictEqual(apps.length, 1);
        assert.strictEqual(firstApp, OfficeApp.Excel);
      });
      it("two apps", function () {
        const apps = parseOfficeApps("OneNote,PowerPoint");
        const [firstApp, secondApp] = apps;
        assert.strictEqual(apps.length, 2);
        assert.strictEqual(firstApp, OfficeApp.OneNote);
        assert.strictEqual(secondApp, OfficeApp.PowerPoint);
      });
      it("two apps with whitespace", function () {
        const apps = parseOfficeApps(" OneNote, PowerPoint ");
        const [firstApp, secondApp] = apps;
        assert.strictEqual(apps.length, 2);
        assert.strictEqual(firstApp, OfficeApp.OneNote);
        assert.strictEqual(secondApp, OfficeApp.PowerPoint);
      });
      it("all", function () {
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
      it("unknown app", function () {
        const unknown = "unknown";
        assert.throws(
          () => {
            parseOfficeApps(unknown);
          },
          new Error(`"${unknown}" is not a valid Office app.`)
        );
        assert.throws(
          () => {
            parseOfficeApps(`Excel,${unknown}`);
          },
          new Error(`"${unknown}" is not a valid Office app.`)
        );
      });
      it("empty string", function () {
        assert.throws(() => {
          parseOfficeApps("");
        }, new Error(`"" is not a valid Office app.`));
      });
    });
    describe("toOfficeApp()", function () {
      it("Excel", function () {
        assert.strictEqual(toOfficeApp("Excel"), OfficeApp.Excel);
        assert.strictEqual(toOfficeApp("excel"), OfficeApp.Excel);
        assert.strictEqual(toOfficeApp("EXCEL"), OfficeApp.Excel);
      });
      it("OneNote", function () {
        assert.strictEqual(toOfficeApp("OneNote"), OfficeApp.OneNote);
        assert.strictEqual(toOfficeApp("onenote"), OfficeApp.OneNote);
        assert.strictEqual(toOfficeApp("ONENOTE"), OfficeApp.OneNote);
      });
      it("Outlook", function () {
        assert.strictEqual(toOfficeApp("Outlook"), OfficeApp.Outlook);
        assert.strictEqual(toOfficeApp("outlook"), OfficeApp.Outlook);
        assert.strictEqual(toOfficeApp("OUTLOOK"), OfficeApp.Outlook);
      });
      it("PowerPoint", function () {
        assert.strictEqual(toOfficeApp("PowerPoint"), OfficeApp.PowerPoint);
        assert.strictEqual(toOfficeApp("powerpoint"), OfficeApp.PowerPoint);
        assert.strictEqual(toOfficeApp("POWERPOINT"), OfficeApp.PowerPoint);
      });
      it("Project", function () {
        assert.strictEqual(toOfficeApp("Project"), OfficeApp.Project);
        assert.strictEqual(toOfficeApp("project"), OfficeApp.Project);
        assert.strictEqual(toOfficeApp("PROJECT"), OfficeApp.Project);
      });
      it("Word", function () {
        assert.strictEqual(toOfficeApp("Word"), OfficeApp.Word);
        assert.strictEqual(toOfficeApp("word"), OfficeApp.Word);
        assert.strictEqual(toOfficeApp("WORD"), OfficeApp.Word);
      });
      it("should return undefined for an unknown or empty value", function () {
        assert.strictEqual(toOfficeApp("unknown"), undefined);
        assert.strictEqual(toOfficeApp(""), undefined);
      });
      it("should trim whitespace", function () {
        assert.strictEqual(toOfficeApp(" excel"), OfficeApp.Excel);
        assert.strictEqual(toOfficeApp("word\n"), OfficeApp.Word);
        assert.strictEqual(toOfficeApp("  \toutlook  "), OfficeApp.Outlook);
        assert.strictEqual(toOfficeApp("  unknown  "), undefined);
        assert.strictEqual(toOfficeApp("    "), undefined);
      });
    });
  });
  describe("manifestInfo.ts", function () {
    describe("readManifestFile() XML", function () {
      it("should read the manifest xml info", async function () {
        const info = await OfficeAddinManifest.readManifestFile(
          "test/manifests/TaskPane.manifest.xml"
        );

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
          assert.strictEqual(
            info.defaultSettings.sourceLocation,
            "https://localhost:3000/taskpane.html"
          );
        } else {
          assert.fail("ManifestInfo should include defaultSettings.");
        }
      });
      it("should throw an error if there is a bad xml end tag", async function () {
        let result;
        try {
          await OfficeAddinManifest.readManifestFile(
            "test/manifests/invalid/incorrect-end-tag.manifest.xml"
          );
        } catch (err: any) {
          result = err;
        }
        assert.equal(
          result.message,
          "Unable to parse the manifest file: test/manifests/invalid/incorrect-end-tag.manifest.xml. \nError: Unexpected close tag\nLine: 8\nColumn: 46\nChar: >"
        );
      });
      it("should handle OfficeApp with no info in xml", async function () {
        const info = await OfficeAddinManifest.readManifestFile(
          "test/manifests/invalid/officeapp-empty.manifest.xml"
        );

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
      it("should handle a missing description in xml", async function () {
        const info = await OfficeAddinManifest.readManifestFile(
          "test/manifests/invalid/no-description.manifest.xml"
        );

        assert.strictEqual(info.defaultLocale, "en-US");
        assert.strictEqual(info.description, undefined);
        assert.strictEqual(info.displayName, "Office Add-in Name");
        assert.strictEqual(info.id, "132a8a21-011a-4ceb-9336-6af8a276a288");
        assert.strictEqual(info.officeAppType, "TaskPaneApp");
        assert.strictEqual(info.providerName, "ProviderName");
        assert.strictEqual(info.version, "1.2.3.4");
      });
    });
    describe("readManifestFile() JSON", function () {
      it("should read the manifest json info", async function () {
        const info: ManifestInfo = await OfficeAddinManifest.readManifestFile(
          path.normalize("test/manifests/manifest.json")
        );
        assert.strictEqual(info.id, "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee");
        assert.strictEqual(info.appDomains instanceof Array, true, "appDomains");
        assert.strictEqual(info.appDomains!.length, 1);
        assert.strictEqual(info.appDomains![0], "contoso.com");
        assert.strictEqual(info.defaultLocale, "en-us");
        assert.strictEqual(info.description, "A template to get started.");
        assert.strictEqual(info.displayName, "Contoso Task Pane Add-in");
        assert.strictEqual(info.highResolutionIconUrl, "assets/color.png", "highResolutionIconUrl");
        assert.strictEqual(info.hosts instanceof Array, true, "hosts");
        assert.strictEqual(info.hosts!.length, 1);
        assert.strictEqual(info.hosts![0], "mail");
        assert.strictEqual(info.iconUrl, "assets/color.png", "iconUrl");
        assert.strictEqual(info.officeAppType, "TaskPaneApp");
        assert.strictEqual(info.permissions, "MailboxItem.Read.User");
        assert.strictEqual(info.providerName, "Contoso");
        assert.strictEqual(info.supportUrl, "https://www.contoso.com");
        assert.strictEqual(info.version, "1.0.0");
      });
      it("should throw an error on an invalid json format", async function () {
        const invalidManifest = path.normalize("test/manifests/invalid/invalid-manifest.json");
        let result;
        try {
          await OfficeAddinManifest.readManifestFile(invalidManifest);
        } catch (err: any) {
          result = err.message;
        }
        assert(
          result.startsWith(
            `Unable to read data for manifest file: ${invalidManifest}. \nSyntaxError: `
          )
        );
      });
      it("nonexistent manifest", async function () {
        const invalidManifest = path.normalize(`${manifestTestFolder}/foo/manifest.json`);
        let result;
        try {
          await OfficeAddinManifest.readManifestFile(invalidManifest);
        } catch (err: any) {
          result = err.message;
        }
        assert.strictEqual(
          result,
          `Unable to read data for manifest file: ${invalidManifest}. \nError: ENOENT: no such file or directory, open '${invalidManifest}'`
        );
      });
      it("invalid extension manifest", async function () {
        this.timeout(6000);
        let result: string = "";
        try {
          await OfficeAddinManifest.readManifestFile("test/foo/tag.manifest.txt");
        } catch (err: any) {
          result = err.message;
        }
        assert.strictEqual(
          result,
          `Manifest operations are not supported in .txt.\nThey are only supported in .xml and in .json.`
        );
      });
    });
    describe("modifyManifestFile() XML", function () {
      beforeEach(async function () {
        await _createManifestTestFolder(manifestTestFolder);
      });
      afterEach(async function () {
        await _deleteFolder(manifestTestFolder);
      });
      it("should handle a specified valid guid and displayName", async function () {
        // call modify, specifying guid and displayName  parameters
        const testGuid = uuidv1();
        const testDisplayName = "TestDisplayName";
        const updatedInfo = await OfficeAddinManifest.modifyManifestFile(
          testManifestXml,
          testGuid,
          testDisplayName
        );

        // verify guid displayName updated
        assert.strictEqual(updatedInfo.id, testGuid);
        assert.strictEqual(updatedInfo.displayName, testDisplayName);
      });
      it(`should handle specifying "random" form guid parameter`, async function () {
        // get original manifest info and create copy of manifest that we can overwrite in this test
        const originalInfo = await OfficeAddinManifest.readManifestFile(testManifestXml);

        // call modify, specifying "random" parameter
        const updatedInfo = await OfficeAddinManifest.modifyManifestFile(
          testManifestXml,
          "random",
          undefined
        );

        // verify guid updated, that it"s a valid guid and that the displayName is not updated
        assert.notStrictEqual(updatedInfo.id, originalInfo.id);
        assert.strictEqual(updatedInfo.id && isUUID(updatedInfo.id), true);
        assert.strictEqual(updatedInfo.displayName, originalInfo.displayName);
      });
      it("should handle specifying displayName only", async function () {
        // get original manifest info and create copy of manifest that we can overwrite in this test
        const originalInfo = await OfficeAddinManifest.readManifestFile(testManifestXml);

        // call  modify, specifying a displayName parameter
        const testDisplayName = "TestDisplayName";
        const updatedInfo = await OfficeAddinManifest.modifyManifestFile(
          testManifestXml,
          undefined,
          testDisplayName
        );

        // verify displayName updated and guid not updated
        assert.notStrictEqual(updatedInfo.displayName, originalInfo.displayName);
        assert.strictEqual(updatedInfo.displayName, testDisplayName);
        assert.strictEqual(updatedInfo.id, originalInfo.id);
      });
      it("should handle not specifying either a guid or displayName", async function () {
        let result;
        try {
          await OfficeAddinManifest.modifyManifestFile(testManifestXml);
        } catch (err: any) {
          result = err.message;
        }
        assert.strictEqual(result, `You need to specify something to change in the manifest.`);
      });
      it("should handle an invalid XML manifest file path", async function () {
        // call  modify, specifying an invalid manifest path with a valid guid and displayName
        const invalidManifest = path.normalize(`${manifestTestFolder}/foo/manifest.xml`);
        const testGuid = uuidv1();
        const testDisplayName = "TestDisplayName";
        let result;
        try {
          await OfficeAddinManifest.modifyManifestFile(invalidManifest, testGuid, testDisplayName);
        } catch (err: any) {
          result = err.message;
        }

        assert.strictEqual(
          result,
          `Unable to modify xml data for manifest file: ${invalidManifest}.\nError: Unable to read data for manifest file: ${invalidManifest}.\nError: ENOENT: no such file or directory, open '${invalidManifest}'`
        );
      });
    });
    describe("modifyManifestFile() JSON", function () {
      beforeEach(async function () {
        await _createManifestTestFolder(manifestTestFolder);
      });
      afterEach(async function () {
        await _deleteFolder(manifestTestFolder);
      });
      it("should handle a specified valid guid and displayName", async function () {
        // call modify, specifying guid and displayName  parameters
        const testGuid = uuidv1();
        const testDisplayName = "TestDisplayName";
        const updatedInfo = await OfficeAddinManifest.modifyManifestFile(
          testManifestJson,
          testGuid,
          testDisplayName
        );

        // verify guid displayName updated
        assert.strictEqual(updatedInfo.id, testGuid);
        assert.strictEqual(updatedInfo.displayName, testDisplayName);
      });
      it(`should handle specifying "random" form guid parameter`, async function () {
        // get original manifest info and create copy of manifest that we can overwrite in this test
        const originalInfo = await OfficeAddinManifest.readManifestFile(testManifestJson);

        // call modify, specifying "random" parameter
        const updatedInfo = await OfficeAddinManifest.modifyManifestFile(
          testManifestJson,
          "random",
          undefined
        );

        // verify guid updated, that it"s a valid guid and that the displayName is not updated
        assert.notStrictEqual(updatedInfo.id, originalInfo.id);
        assert.strictEqual(updatedInfo.id && isUUID(updatedInfo.id), true);
        assert.strictEqual(updatedInfo.displayName, originalInfo.displayName);
      });
      it("should handle specifying displayName only", async function () {
        // get original manifest info and create copy of manifest that we can overwrite in this test
        const originalInfo = await OfficeAddinManifest.readManifestFile(testManifestJson);

        // call  modify, specifying a displayName parameter
        const testDisplayName = "TestDisplayName";
        const updatedInfo = await OfficeAddinManifest.modifyManifestFile(
          testManifestJson,
          undefined,
          testDisplayName
        );

        // verify displayName updated and guid not updated
        assert.notStrictEqual(updatedInfo.displayName, originalInfo.displayName);
        assert.strictEqual(updatedInfo.displayName, testDisplayName);
        assert.strictEqual(updatedInfo.id, originalInfo.id);
      });
      it("should handle not specifying either a guid or displayName", async function () {
        let result;
        try {
          await OfficeAddinManifest.modifyManifestFile(testManifestJson);
        } catch (err: any) {
          result = err.message;
        }
        assert.strictEqual(result, `You need to specify something to change in the manifest.`);
      });
      it("should handle an invalid JSON manifest file path", async function () {
        // call  modify, specifying an invalid manifest path with a valid guid and displayName
        const invalidManifest = path.normalize(`${manifestTestFolder}/foo/manifest.json`);
        const testGuid = uuidv1();
        const testDisplayName = "TestDisplayName";
        let result;
        try {
          await OfficeAddinManifest.modifyManifestFile(invalidManifest, testGuid, testDisplayName);
        } catch (err: any) {
          result = err.message;
        }

        assert.strictEqual(
          result,
          `Unable to modify json data for manifest file: ${invalidManifest}. \nError: ENOENT: no such file or directory, open '${invalidManifest}'`
        );
      });
    });
  });
  describe("validate.ts", function () {
    describe("validateManifest() XML", function () {
      this.slow(5000);
      it("valid manifest", async function () {
        this.timeout(6000);
        const validation = await validateManifest("test/manifests/TaskPane.manifest.xml");
        assert.strictEqual(validation.isValid, true);
        assert.strictEqual(validation.status, 200);
        assert.strictEqual(validation.report!.errors!.length, 0);
        assert.strictEqual(validation.report!.notes!.length > 0, true);
        assert.strictEqual(validation.report!.warnings!.length, 0);
        assert.strictEqual(validation.report!.addInDetails!.supportedProducts!.length > 0, true);
      });
      it("valid manifest prod (fail)", async function () {
        this.timeout(6000);
        const validation = await validateManifest("test/manifests/TaskPane.manifest.xml", true);
        assert.strictEqual(validation.isValid, false);
        assert.strictEqual(validation.status, 200);
        assert.strictEqual(validation.report!.errors!.length, 6);
        assert.strictEqual(validation.report!.notes!.length > 0, true);
        assert.strictEqual(validation.report!.warnings!.length, 0);
        assert.strictEqual(validation.report!.addInDetails!.supportedProducts!.length > 0, true);
      });
      it("invalid manifest path", async function () {
        this.timeout(6000);
        let result: string = "";
        const invalidManifestPath = path.normalize(`${manifestTestFolder}/foo/manifest.xml`);
        try {
          await validateManifest(invalidManifestPath);
        } catch (err: any) {
          result = err.message;
        }
        assert.strictEqual(result.indexOf("ENOENT: no such file or directory") >= 0, true);
      });
      it("Excel", async function () {
        this.timeout(6000);
        const validation = await validateManifest("test/manifests/TaskPane.Excel.manifest.xml");
        assert.strictEqual(validation.isValid, true);
        assert.strictEqual(validation.status, 200);
        assert.strictEqual(validation.report!.errors!.length, 0);
        assert.strictEqual(validation.report!.notes!.length > 0, true);
        assert.strictEqual(validation.report!.warnings!.length, 0);
        assert.strictEqual(validation.report!.addInDetails!.supportedProducts!.length > 0, true);
      });
      it("OneNote", async function () {
        this.timeout(6000);
        const validation = await validateManifest("test/manifests/TaskPane.OneNote.manifest.xml");
        assert.strictEqual(validation.isValid, true);
        assert.strictEqual(validation.status, 200);
        assert.strictEqual(validation.report!.errors!.length, 0);
        assert.strictEqual(validation.report!.notes!.length > 0, true);
        assert.strictEqual(validation.report!.warnings!.length, 0);
        assert.strictEqual(validation.report!.addInDetails!.supportedProducts!.length > 0, true);
      });
      it("Outlook", async function () {
        this.timeout(6000);
        const validation = await validateManifest("test/manifests/TaskPane.Outlook.manifest.xml");
        assert.strictEqual(validation.isValid, true);
        assert.strictEqual(validation.status, 200);
        assert.strictEqual(validation.report!.errors!.length, 0);
        assert.strictEqual(validation.report!.notes!.length > 0, true);
        assert.strictEqual(validation.report!.warnings!.length, 0);
        assert.strictEqual(validation.report!.addInDetails!.supportedProducts!.length > 0, true);
      });
      it("PowerPoint", async function () {
        this.timeout(6000);
        const validation = await validateManifest(
          "test/manifests/TaskPane.PowerPoint.manifest.xml"
        );
        assert.strictEqual(validation.isValid, true);
        assert.strictEqual(validation.status, 200);
        assert.strictEqual(validation.report!.errors!.length, 0);
        assert.strictEqual(validation.report!.notes!.length > 0, true);
        assert.strictEqual(validation.report!.warnings!.length, 0);
        assert.strictEqual(validation.report!.addInDetails!.supportedProducts!.length > 0, true);
      });
      it("Project", async function () {
        const validation = await validateManifest("test/manifests/TaskPane.Project.manifest.xml");
        assert.strictEqual(validation.isValid, true);
        assert.strictEqual(validation.status, 200);
        assert.strictEqual(validation.report!.errors!.length, 0);
        assert.strictEqual(validation.report!.notes!.length > 0, true);
        assert.strictEqual(validation.report!.warnings!.length, 0);
        assert.strictEqual(validation.report!.addInDetails!.supportedProducts!.length > 0, true);
      });
      it("Word", async function () {
        this.timeout(6000);
        const validation = await validateManifest("test/manifests/TaskPane.Word.manifest.xml");
        assert.strictEqual(validation.isValid, true);
        assert.strictEqual(validation.status, 200);
        assert.strictEqual(validation.report!.errors!.length, 0);
        assert.strictEqual(validation.report!.notes!.length > 0, true);
        assert.strictEqual(validation.report!.warnings!.length, 0);
        assert.strictEqual(validation.report!.addInDetails!.supportedProducts!.length > 0, true);
      });
    });

    describe("validateManifest() JSON", function () {
      this.slow(5000);
      it("Valid teams manifest", async function () {
        this.timeout(6000);
        const validation = await validateManifest("test/manifests/teamsManifest.json");
        assert.strictEqual(validation.isValid, true);
      });
      it("invalid manifest path", async function () {
        this.timeout(6000);
        let result: string = "";
        const invalidManifestPath = path.normalize(`${manifestTestFolder}/foo/manifest.json`);
        try {
          await validateManifest(invalidManifestPath);
        } catch (err: any) {
          result = err.message;
        }
        assert.strictEqual(result.indexOf("ENOENT: no such file or directory") >= 0, true);
      });
    });
  });
  describe("export.ts", function () {
    describe("exportMetadataPackage()", function () {
      it("export manifest to test location", async function () {
        this.timeout(6000);
        const testFolder = path.resolve("./testExecution");
        const manifestPath = path.normalize("test/manifests/manifest.json");
        const outputPath = path.normalize(`${testFolder}/testPackage.zip`);
        const outputFile = await exportMetadataPackage(outputPath, manifestPath);

        assert.strictEqual(
          outputFile,
          outputPath,
          "Output path '" + outputFile + "' should match the argument '" + outputPath + "'"
        );
        assert.strictEqual(
          fs.existsSync(outputFile),
          true,
          "Output file '" + outputFile + "' should exist"
        );

        // Cleanup
        fs.rmSync(testFolder, { recursive: true });
      });
      it("export manifest to default location", async function () {
        this.timeout(6000);
        const manifestPath = path.normalize("test/manifests/manifest.json");
        const expectedOutput = path.join(path.dirname(path.resolve(manifestPath)), "manifest.zip");
        const outputFile = await exportMetadataPackage("", manifestPath);
        assert.strictEqual(
          outputFile,
          expectedOutput,
          "Output path '" + outputFile + "' should match the default '" + expectedOutput + "'"
        );
        assert.strictEqual(
          fs.existsSync(outputFile),
          true,
          "Output file '" + outputFile + "' should exist"
        );

        // Cleanup
        fs.unlinkSync(outputFile);
      });
      it("export manifest with different name", async function () {
        this.timeout(6000);
        const manifestPath = path.normalize("test/manifests/manifest.local.json");
        const expectedOutput = path.join(path.dirname(path.resolve(manifestPath)), "manifest.zip");
        const outputFile = await exportMetadataPackage("", manifestPath);
        assert.strictEqual(
          outputFile,
          expectedOutput,
          "Output path '" + outputFile + "' should match the default '" + expectedOutput + "'"
        );
        assert.strictEqual(
          fs.existsSync(outputFile),
          true,
          "Output file '" + outputFile + "' should exist"
        );

        const zip: AdmZip = new AdmZip(outputFile);
        const entries = zip.getEntries().filter((entry) => {
          return entry.name == "manifest.json";
        });
        assert.strictEqual(entries.length, 1, "Found manifest.json in zip file");
        // Cleanup
        fs.unlinkSync(outputFile);
      });
      it("invalid manifest path", async function () {
        this.timeout(6000);
        const invalidManifestPath = path.normalize(`${manifestTestFolder}/foo/manifest.json`);
        const expectedError = `The file '${invalidManifestPath}' does not exist`;
        let result: string = "";

        try {
          await exportMetadataPackage("" /* use default output path */, invalidManifestPath);
        } catch (err: any) {
          result = err.message;
        }
        assert.strictEqual(result, expectedError);
      });
    });
  });
});

async function _deleteFolder(folder: string): Promise<void> {
  if (fs.existsSync(folder)) {
    fs.readdirSync(folder).forEach(function (file) {
      const curPath = path.join(folder, file);

      if (fs.lstatSync(curPath).isDirectory()) {
        _deleteFolder(curPath);
      } else {
        fs.unlinkSync(curPath);
      }
    });
    fs.rmdirSync(folder);
  }
}

async function _createManifestTestFolder(folder: string): Promise<void> {
  if (fs.existsSync(folder)) {
    await _deleteFolder(folder);
  }
  const fsExtra = require("fs-extra");
  await fsExtra.copy(manifestOriginalFolder, folder);
}
