import * as assert from "assert";
import * as fs from "fs";
import * as mocha from "mocha";
import * as uuid from "uuid";
import { isUUID } from "validator";
import * as manifestInfo from "../src/manifestInfo";

const manifestOriginalFolder = process.cwd() + "/test/manifests";
const manifestTestFolder = process.cwd() + "\\testExecution\\testManifests";
const testManifest = manifestTestFolder + "\\Taskpane.manifest.xml";

describe("Manifest", function() {
  describe("readManifestInfo", function() {
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
});

describe("Manifest", function() {
  this.beforeEach(async function() {
    await _createManifestFilesFolder();
  });
  this.afterEach(async function() {
    await _deleteManifestTestFolder(manifestTestFolder);
  });
  describe("modifyManifestFile", function() {
    it("should handle a specified valid guid and displayName", async function() {
      // call modify, specifying guid and displayName  parameters
      const testGuid = uuid.v1();
      const testDisplayName = "TestDisplayName";
      const updatedInfo = await manifestInfo.modifyManifestFile(testManifest, testGuid, testDisplayName);

      // verify guid displayName updated
      assert.strictEqual(updatedInfo.id, testGuid);
      assert.strictEqual(updatedInfo.displayName, testDisplayName);
    });
    it("should handle specifying \"random\" form guid parameter", async function() {
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
      const invalidManifest = manifestTestFolder + "\\foo\\manifest.xml";
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
