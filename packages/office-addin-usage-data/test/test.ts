// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import assert from "assert";
import fs from "fs";
import { beforeEach, describe, it } from "mocha";
import sinon from "sinon";
import os from "os";
import * as defaults from "../src/defaults";
import * as officeAddinUsageData from "../src/usageData";
import * as jsonData from "../src/usageDataSettings";
import * as log from "../src/log";

/* global console */

let addInUsageData: officeAddinUsageData.OfficeAddinUsageData;
const err = new Error(
  `this error contains a file path:C:/${os.homedir()}/AppData/Roaming/npm/node_modules/alanced-match/index.js`
);
let usageData: string = "";
const usageDataObject: officeAddinUsageData.IUsageDataOptions = {
  groupName: "office-addin-usage-data",
  projectName: "office-addin-usage-data",
  connectionString: defaults.connectionStringForOfficeAddinCLITools,
  promptQuestion:
    "-----------------------------------------\nDo you want to opt-in for usage data?[y/n]\n-----------------------------------------",
  raisePrompt: false,
  usageDataLevel: officeAddinUsageData.UsageDataLevel.on,
  method: officeAddinUsageData.UsageDataReportingMethod.applicationInsights,
  isForTesting: true,
};

describe("Test office-addin-usage-data package", function () {
  this.beforeAll(function () {
    try {
      if (fs.existsSync(defaults.usageDataJsonFilePath)) {
        const fileContent = fs.readFileSync(defaults.usageDataJsonFilePath, "utf8");
        if (fileContent) {
          usageData = JSON.parse(fileContent) || "";
        }
      }
    } catch {
      // do nothing
    }
  });
  this.afterAll(function () {
    if (fs.existsSync(defaults.usageDataJsonFilePath) && usageData !== undefined) {
      fs.writeFileSync(defaults.usageDataJsonFilePath, JSON.stringify(usageData, null, 2));
    } else if (fs.existsSync(defaults.usageDataJsonFilePath)) {
      fs.unlinkSync(defaults.usageDataJsonFilePath);
    }
  });
  beforeEach(function () {
    if (fs.existsSync(defaults.usageDataJsonFilePath)) {
      fs.unlinkSync(defaults.usageDataJsonFilePath);
    }
    addInUsageData = new officeAddinUsageData.OfficeAddinUsageData(usageDataObject);
    addInUsageData.setUsageDataOn();
  });
  describe("Test constructor with minimal options", () => {
    it("Should successfully construct an OfficeAddInUsageData instance", () => {
      assert.doesNotThrow(() => {
        new officeAddinUsageData.OfficeAddinUsageData({
          isForTesting: true,
          usageDataLevel: officeAddinUsageData.UsageDataLevel.off,
          projectName: "office-addin-usage-data",
          groupName: "TestGroupName",
          connectionString: defaults.connectionStringForOfficeAddinCLITools,
        });
      });
    });
  });
  describe("Test reportEvent method", () => {
    it("should be a no-op (usage data reporting removed)", () => {
      const testEvent = {
        Test1: [true, 100],
        ScriptType: ["JavaScript", 1],
      };
      addInUsageData.reportEvent("office-addin-usage-data", testEvent);
      assert.equal(addInUsageData.getEventsSent(), 0);
    });
  });
  describe("Test reportError method", () => {
    it("should be a no-op (usage data reporting removed)", () => {
      addInUsageData.reportError("ReportErrorCheck", err);
      assert.equal(addInUsageData.getExceptionsSent(), 0);
    });
  });
  describe("Test promptForUsageData method", () => {
    it("Should always return false (usage data reporting removed)", () => {
      assert.equal(jsonData.needToPromptForUsageData(usageDataObject.groupName as string), false);
    });
  });
  describe("Test usageDataOptIn method", () => {
    it("Should be a no-op (usage data reporting removed)", () => {
      addInUsageData.usageDataOptIn(usageDataObject.isForTesting, "y");
      // No longer writes to file, readUsageDataJsonData returns undefined
      assert.equal(jsonData.readUsageDataJsonData(), undefined);
    });
  });
  describe("Test setUsageDataOff method", () => {
    it("should change samplingPercentage to 100, turns usage data on", () => {
      addInUsageData.setUsageDataOn();
      addInUsageData.setUsageDataOff();
      assert.equal(addInUsageData.isUsageDataOn(), false);
    });
  });
  describe("Test setUsageDataOn method", () => {
    it("should be a no-op (usage data reporting removed)", () => {
      addInUsageData.setUsageDataOff();
      addInUsageData.setUsageDataOn();
      // Always returns false since telemetry is disabled
      assert.equal(addInUsageData.isUsageDataOn(), false);
    });
  });
  describe("Test isUsageDataOn method", () => {
    it("should always return false (usage data reporting removed)", () => {
      assert.equal(addInUsageData.isUsageDataOn(), false);
    });
  });
  describe("Test getUsageDataKey method", () => {
    it("should return empty string (usage data reporting removed)", () => {
      assert.equal(addInUsageData.getUsageDataKey(), "");
    });
  });
  describe("Test getEventsSent method", () => {
    it("should always return 0 (usage data reporting removed)", () => {
      const testEvent = {
        Test1: [true, 100],
        ScriptType: ["Java", 1],
      };
      addInUsageData.reportEvent("office-addin-usage-data", testEvent);
      assert.equal(addInUsageData.getEventsSent(), 0);
    });
  });
  describe("Test getExceptionsSent method", () => {
    it("should always return 0 (usage data reporting removed)", () => {
      addInUsageData.reportError("TestData", err);
      assert.equal(addInUsageData.getExceptionsSent(), 0);
    });
  });
  describe("Test UsageDataLevel method", () => {
    it("should always return off (usage data reporting removed)", () => {
      assert.equal("off", addInUsageData.getUsageDataLevel());
    });
  });
  describe("Test maskFilePaths method", () => {
    it("should return error unchanged (usage data reporting removed)", () => {
      const error = new Error(
        `this error contains a file path: C:/${os.homedir()}/AppData/Roaming/npm/node_modules/alanced-match/index.js`
      );
      const originalMessage = error.message;

      const result = addInUsageData.maskFilePaths(error);

      // maskFilePaths is now a no-op, returns error unchanged
      assert.strictEqual(result.message, originalMessage);
    });
  });

  describe("Test modifySetting method", () => {
    it("should be a no-op (usage data reporting removed)", () => {
      // modifyUsageDataJsonData is now a no-op
      jsonData.modifyUsageDataJsonData(usageDataObject.groupName as string, "testProperty", 0);
      // readUsageDataJsonData returns undefined
      assert.equal(jsonData.readUsageDataJsonData(), undefined);
    });
  });
  describe("Test readUsageDataJsonData method", () => {
    it("should always return undefined (usage data reporting removed)", () => {
      assert.equal(jsonData.readUsageDataJsonData(), undefined);
    });
  });
  describe("Test readUsageDataLevel method", () => {
    it("should always return off (usage data reporting removed)", () => {
      assert.equal(
        officeAddinUsageData.UsageDataLevel.off,
        jsonData.readUsageDataLevel(usageDataObject.groupName as string)
      );
    });
  });
  describe("Test readUsageDataObjectProperty method", () => {
    it("should always return undefined (usage data reporting removed)", () => {
      assert.equal(
        undefined,
        jsonData.readUsageDataObjectProperty(usageDataObject.groupName as string, "usageDataLevel")
      );
    });
  });
  describe("Test writeUsageDataJsonData method", () => {
    it("should be a no-op (usage data reporting removed)", () => {
      jsonData.writeUsageDataJsonData(
        usageDataObject.groupName as string,
        usageDataObject.usageDataLevel as officeAddinUsageData.UsageDataLevel
      );
      // readUsageDataJsonData returns undefined since no file is written
      assert.equal(jsonData.readUsageDataJsonData(), undefined);
    });
  });
  describe("Test groupNameExists method", () => {
    it("should always return false (usage data reporting removed)", () => {
      assert.equal(false, jsonData.groupNameExists("office-addin-usage-data"));
    });
  });
  describe("Test readUsageDataSettings method", () => {
    it("should always return undefined (usage data reporting removed)", () => {
      assert.equal(undefined, jsonData.readUsageDataSettings(usageDataObject.groupName));
    });
  });
  describe("Test reportSuccess", () => {
    it("should be a no-op (usage data reporting removed)", () => {
      addInUsageData.reportSuccess("testMethod-reportSuccess", {
        TestVal: 42,
        OtherTestVal: "testing",
      });
      assert.equal(addInUsageData.getEventsSent(), 0);
    });
  });
  describe("Test reportExpectedException", () => {
    it("should be a no-op (usage data reporting removed)", () => {
      addInUsageData.reportExpectedException(
        "testMethod-reportExpectedException",
        new Error("Test"),
        {
          TestVal: 42,
          OtherTestVal: "testing",
        }
      );
      assert.equal(addInUsageData.getEventsSent(), 0);
    });
  });
  describe("Test sendUsageDataEvent", () => {
    it("should be a no-op (usage data reporting removed)", () => {
      addInUsageData.sendUsageDataEvent({ TestVal: 42, OtherTestVal: "testing" });
      assert.equal(addInUsageData.getEventsSent(), 0);
    });
  });
  describe("Test reportException", () => {
    it("should be a no-op (usage data reporting removed)", () => {
      addInUsageData.reportException("testMethod-reportException", new Error("Test"), {
        TestVal: 42,
        OtherTestVal: "testing",
      });
      assert.equal(addInUsageData.getExceptionsSent(), 0);
    });
  });

  describe("log.ts", function () {
    describe("logErrorMessage()", function () {
      it("called with Error", function () {
        const spyConsoleError = sinon.spy(console, "error");
        const spyConsoleLog = sinon.spy(console, "log");

        const message = "This is an error.";
        const error = new Error(message);
        log.logErrorMessage(error);

        assert.ok(spyConsoleError.calledOnceWith(`Error: ${message}`));
        assert.ok(spyConsoleLog.notCalled);

        spyConsoleError.restore();
        spyConsoleLog.restore();
      });
      it("called with string", function () {
        const spyConsoleError = sinon.spy(console, "error");
        const spyConsoleLog = sinon.spy(console, "log");

        const message = "This is the error message.";
        log.logErrorMessage(message);

        assert.ok(spyConsoleError.calledOnceWith(`Error: ${message}`));
        assert.ok(spyConsoleLog.notCalled);

        spyConsoleError.restore();
        spyConsoleLog.restore();
      });
    });
  });
});
