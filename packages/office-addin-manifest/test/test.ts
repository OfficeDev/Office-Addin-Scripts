import * as assert from "assert";
import * as fs from "fs";
import * as mocha from "mocha";
import * as manifestInfo from "../src/manifestInfo";
const uuid = require('uuid');
const validator = require("validator");
const manifestOriginalFolder = process.cwd() + "/test/manifests";
const manifestTestFolder = process.cwd() + "\\testExecution\\testManifests";
const testManifest = manifestTestFolder + "\\manifest.xml";

describe("Manifest", function() {
  describe("readManifestInfo", function() {
    it("should read the manifest info", async function() {
      const info = await manifestInfo.readManifestFile("test/manifests/manifest.xml");

      assert.strictEqual(info.defaultLocale, "en-US");
      assert.strictEqual(info.description, "Describes this Office Add-in.");
      assert.strictEqual(info.displayName, "Office Add-in Name");
      assert.strictEqual(info.id, "132a8a21-011a-4ceb-9336-6af8a276a288");
      assert.strictEqual(info.officeAppType, "TaskPaneApp");
      assert.strictEqual(info.providerName, "ProviderName");
      assert.strictEqual(info.version, "1.2.3.4");
    });
    it("should throw an error if there is a bad xml end tag", async function() {
        let result;
        try {
          const info = await manifestInfo.readManifestFile("test/manifests/manifest.incorrect-end-tag.xml");
        } catch (err) {
          result = err;
        }
        assert.equal(result, "Unable to parse the manifest file: test/manifests/manifest.incorrect-end-tag.xml. \nError: Unexpected close tag\nLine: 8\nColumn: 46\nChar: >");        
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
      assert.notStrictEqual(originalInfo.id, updatedInfo.id);
      assert.equal(true, validator.isUUID(updatedInfo.id));
      assert.strictEqual(originalInfo.displayName, updatedInfo.displayName);
    });
    it("should handle specifying displayName only", async function() {
      // get original manifest info and create copy of manifest that we can overwrite in this test
      const originalInfo = await manifestInfo.readManifestFile(testManifest);

      // call  modify, specifying a displayName parameter
      const testDisplayName = "TestDisplayName";
      const updatedInfo = await manifestInfo.modifyManifestFile(testManifest, undefined, testDisplayName);

      // verify displayName updated and guid not updated
      assert.notStrictEqual(originalInfo.displayName, updatedInfo.displayName);
      assert.strictEqual(originalInfo.id, updatedInfo.id);
    });
    it("should handle not specifying either a guid or displayName", async function() {
      const result =  await manifestInfo.modifyManifestFile(testManifest, undefined, undefined);
      assert.equal(result, "Error: You need to specify something to change in the manifest.");
    });
    it("should handle an invalid manifest file path", async function() {
      // call  modify, specifying an invalid manifest path with a valid guid and displayName
      const invalidManifest = manifestTestFolder + "\\foo\\manifest.xml";
      const testGuid = uuid.v1();
      const testDisplayName = "TestDisplayName";
      const  result = await manifestInfo.modifyManifestFile(invalidManifest, testGuid, testDisplayName);
      assert.strictEqual(result, `Unable to read the manifest file: ${invalidManifest}. \nError: ENOENT: no such file or directory, open '${invalidManifest}'`);
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
