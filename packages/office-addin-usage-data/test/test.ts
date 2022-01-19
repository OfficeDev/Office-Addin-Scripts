// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as appInsights from "applicationinsights";
import * as assert from "assert";
import * as fs from "fs";
import * as mocha from "mocha";
import * as os from "os";
import * as defaults from "../src/defaults";
import * as officeAddinUsageData from "../src/usageData";
import * as jsonData from "../src/usageDataSettings";

let addInUsageData: officeAddinUsageData.OfficeAddinUsageData;
const err = new Error(`this error contains a file path:C:/Users/admin/AppData/Roaming/npm/node_modules/alanced-match/index.js`);
let usageData: string;
const usageDataObject: officeAddinUsageData.IUsageDataOptions = {
  groupName: "office-addin-usage-data",
  projectName: "office-addin-usage-data",
  instrumentationKey: defaults.instrumentationKeyForOfficeAddinCLITools,
  promptQuestion: "-----------------------------------------\nDo you want to opt-in for usage data?[y/n]\n-----------------------------------------",
  raisePrompt: false,
  usageDataLevel: officeAddinUsageData.UsageDataLevel.on,
  method: officeAddinUsageData.UsageDataReportingMethod.applicationInsights,
  isForTesting: true,
};

describe("Test office-addin-usage data-package", function() {
  this.beforeAll(function() {
    try {
    if (fs.existsSync(defaults.usageDataJsonFilePath) && fs.readFileSync(defaults.usageDataJsonFilePath, "utf8").toString() !== undefined) {
      usageData = JSON.parse(fs.readFileSync(defaults.usageDataJsonFilePath, "utf8").toString());
    }
  } catch {
    usageData = undefined;
  }
  });
  this.afterAll(function() {
    if (fs.existsSync(defaults.usageDataJsonFilePath) && usageData !== undefined) {
      fs.writeFileSync(defaults.usageDataJsonFilePath, JSON.stringify((usageData), null, 2));
    } else if (fs.existsSync(defaults.usageDataJsonFilePath)) {
        fs.unlinkSync(defaults.usageDataJsonFilePath);
    }
  });
  beforeEach(function() {
    if (fs.existsSync(defaults.usageDataJsonFilePath)) {
      fs.unlinkSync(defaults.usageDataJsonFilePath);
    }
    addInUsageData = new officeAddinUsageData.OfficeAddinUsageData(usageDataObject);
    addInUsageData.setUsageDataOn();
  });
  describe("Test constructor with minimal options", () => {
    it("Should successfully construct an OfficeAddInUsageData instance", () => {
      assert.doesNotThrow(() => {
        let testAddInUsageData = new officeAddinUsageData.OfficeAddinUsageData({
          isForTesting: true,
          usageDataLevel: officeAddinUsageData.UsageDataLevel.off,
          projectName: "office-addin-usage-data",
          groupName: "TestGroupName",
          instrumentationKey: defaults.instrumentationKeyForOfficeAddinCLITools
        });
      });
    });
  });
  describe("Test reportEvent method", () => {
    it("should track event of object passed in with a project name", () => {
      const testEvent = {
        Test1: [true, 100],
        ScriptType: ["JavaScript", 1],
      };
      addInUsageData.reportEvent("office-addin-usage-data", testEvent);
      assert.equal(addInUsageData.getEventsSent(), 1);
    });
  });
  describe("Test reportError method", () => {
    it("should send usage data exception", () => {
      addInUsageData.reportError("ReportErrorCheck", err);
      assert.equal(addInUsageData.getExceptionsSent(), 1);
    });
  });
  describe("Test promptForUsageData method", () => {
    it("Should return 'true' because usageDataJsonFilePath doesn't exist", () => {
      // delete officeAddinUsageData.json
      if (fs.existsSync(defaults.usageDataJsonFilePath)) {
        fs.unlinkSync(defaults.usageDataJsonFilePath);
      }
      assert.equal(jsonData.needToPromptForUsageData(usageDataObject.groupName), true);
    });
  });
  describe("Test promptForUsageData method", () => {
    it("Should return 'false' because usageDataJsonFilePath exists and groupName exists in file", () => {
      jsonData.writeUsageDataJsonData(usageDataObject.groupName, usageDataObject.usageDataLevel);
      assert.equal(jsonData.needToPromptForUsageData(usageDataObject.groupName), false);
    });
  });
  describe("Test promptForUsageData method", () => {
    it("Should return 'true' because usageDataJsonFilePath exists but groupName doesn't exist on file", () => {
      jsonData.writeUsageDataJsonData(usageDataObject.groupName, usageDataObject.usageDataLevel);
      assert.equal(jsonData.needToPromptForUsageData("test group name"), true);
    });
  });
  describe("Test usageDataOptIn method", () => {
    it("Should write out file with groupName set to true to usageDataJsonFilePath", () => {
      addInUsageData.usageDataOptIn(usageDataObject.isForTesting, "y");
      const jsonTelemtryData = jsonData.readUsageDataJsonData();
      assert.equal(jsonTelemtryData.usageDataInstances[usageDataObject.groupName].usageDataLevel, usageDataObject.usageDataLevel);
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
    it("should change samplingPercentage to 100, turns usage data on", () => {
      addInUsageData.setUsageDataOff();
      addInUsageData.setUsageDataOn();
      assert.equal(appInsights.defaultClient.config.samplingPercentage, 100);
    });
  });
  describe("Test isUsageDataOn method", () => {
    it("should return true if samplingPercentage is on(100)", () => {
      appInsights.defaultClient.config.samplingPercentage = 100;
      assert.equal(addInUsageData.isUsageDataOn(), true);
    });
    it("should return false if samplingPercentage is off(0)", () => {
      appInsights.defaultClient.config.samplingPercentage = 0;
      assert.equal(addInUsageData.isUsageDataOn(), false);
    });
  });
  describe("Test getUsageDataKey method", () => {
    it("should return usage data key", () => {
      assert.equal(addInUsageData.getUsageDataKey(), "de0d9e7c-1f46-4552-bc21-4e43e489a015");
    });
  });
  describe("Test getEventsSent method", () => {
    it("should return amount of events successfully sent", () => {
      addInUsageData.setUsageDataOff();
      const testEvent = {
        Test1: [true, 100],
        ScriptType: ["Java", 1],
      };
      addInUsageData.reportEvent("office-addin-usage-data", testEvent);
      assert.equal(addInUsageData.getEventsSent(), 1);
    });
  });
  describe("Test getExceptionsSent method", () => {
    it("should return amount of exceptions successfully sent ", () => {
      addInUsageData.setUsageDataOff();
      addInUsageData.reportError("TestData", err);
      assert.equal(addInUsageData.getExceptionsSent(), 1);
    });
  });
  describe("Test UsageDataLevel method", () => {
    it("should return the usage data level of the object", () => {
      assert.equal("on", addInUsageData.getUsageDataLevel());
    });
  });
  describe("Test maskFilePaths method", () => {
    it("should parse error file paths with slashs", () => {
      addInUsageData.setUsageDataOff();
      const error = new Error(`this error contains a file path:C:/Users/admin/AppData/Roaming/npm/node_modules/alanced-match/index.js`);
      const compareError = new Error();
      compareError.name = "Error";
      compareError.message = "this error contains a file path:C:\\.js";
      // may throw error if change any part of the top of the test file
      compareError.stack = "this error contains a file path:C:\\.js";

      addInUsageData.maskFilePaths(error);

      assert.strictEqual(compareError.name, error.name);
      assert.strictEqual(compareError.message, error.message);
      assert.strictEqual(error.stack.includes(compareError.stack), true);
    });
    it("should parse error file paths with backslashs", () => {
      addInUsageData.setUsageDataOff();
      const errWithBackslash = new Error(`this error contains a file path:C:\\Users\\admin\\AppData\\Local\\Temp\\excel file .xlsx`);
      const compareErrorWithBackslash = new Error();
      compareErrorWithBackslash.message = "this error contains a file path:C:\\.xlsx";
      compareErrorWithBackslash.stack = "this error contains a file path:C:\\.xlsx";
      
      addInUsageData.maskFilePaths(errWithBackslash);

      assert.strictEqual(compareErrorWithBackslash.message, errWithBackslash.message);
      assert.strictEqual(errWithBackslash.stack.includes(compareErrorWithBackslash.stack), true);
    });
    it("should parse error file paths with slashs and backslashs", () => {
      addInUsageData.setUsageDataOff();
      const error = new Error(`this error contains a file path:C:\\Users/\\admin\\AppData\\Local//Temp\\excel_file .xlsx`);
      const compareError = new Error();
      compareError.message = "this error contains a file path:C:\\.xlsx";
      compareError.stack = "this error contains a file path:C:\\.xlsx";;
      
      addInUsageData.maskFilePaths(error);

      assert.strictEqual(compareError.message, error.message);
      assert.strictEqual(error.stack.includes(compareError.stack), true);
    });
    it("should handle relative paths", () => {
      addInUsageData.setUsageDataOff();
      const error = new Error(`file path: /this is a file path/path/manifestName.xml`);
      const compareError = new Error();
      compareError.message = "file path: \\.xml";
      compareError.stack = "file path: \\.xml";
      
      addInUsageData.maskFilePaths(error);

      assert.strictEqual(compareError.message, error.message);
      assert.strictEqual(error.stack.includes(compareError.stack), true);
    });
  });

  describe("Test modifySetting method", () => {
    it("should modify or create specific property to new value", () => {
      const usageDataLevel = usageDataObject.usageDataLevel;
      let jsonObject = {};
      jsonObject[usageDataObject.groupName] = usageDataObject.usageDataLevel;
      jsonObject = { usageDataInstances: jsonData };
      jsonObject = { usageDataInstances: { [usageDataObject.groupName]: { usageDataLevel } } };
      fs.writeFileSync(defaults.usageDataJsonFilePath, JSON.stringify((jsonObject)));
      const usageDataJsonData = jsonData.readUsageDataJsonData();
      const testPropertyName = "testProperty";
      usageDataJsonData.usageDataInstances[usageDataObject.groupName][testPropertyName] = 0;
      jsonData.modifyUsageDataJsonData(usageDataObject.groupName, testPropertyName, 0);
      assert.equal(JSON.stringify(usageDataJsonData), JSON.stringify(jsonData.readUsageDataJsonData()));
    });
  });
  describe("Test readUsageDataJsonData method", () => {
    it("should read and return parsed object object from usage data", () => {
      const usageDataLevel = usageDataObject.usageDataLevel;
      let jsonObject = {};
      jsonObject[usageDataObject.groupName] = usageDataObject.usageDataLevel;
      jsonObject = { usageDataInstances: jsonData };
      jsonObject = { usageDataInstances: { [usageDataObject.groupName]: { usageDataLevel } } };
      fs.writeFileSync(defaults.usageDataJsonFilePath, JSON.stringify((jsonObject)));
      assert.equal(JSON.stringify(jsonObject), JSON.stringify(jsonData.readUsageDataJsonData()));
    });
  });
  describe("Test readUsageDataLevel method", () => {
    it("should read and return object's usage data level from file", () => {
      const usageDataLevel = usageDataObject.usageDataLevel;
      let jsonObject = {};
      jsonObject[usageDataObject.groupName] = usageDataObject.usageDataLevel;
      jsonObject = { usageDataInstances: jsonData };
      jsonObject = { usageDataInstances: { [usageDataObject.groupName]: { usageDataLevel } } };
      fs.writeFileSync(defaults.usageDataJsonFilePath, JSON.stringify((jsonObject)));
      assert.equal(officeAddinUsageData.UsageDataLevel.on, jsonData.readUsageDataLevel(usageDataObject.groupName));
    });
  });
  describe("Test readUsageDataObjectProperty method", () => {
    it("should read and return parsed object object from usage data", () => {
      const usageDataLevel = usageDataObject.usageDataLevel;
      let jsonObject = {};
      jsonObject[usageDataObject.groupName] = usageDataObject.usageDataLevel;
      jsonObject = { usageDataInstances: jsonData };
      jsonObject = { usageDataInstances: { [usageDataObject.groupName]: { usageDataLevel } } };
      fs.writeFileSync(defaults.usageDataJsonFilePath, JSON.stringify((jsonObject)));
      assert.equal(officeAddinUsageData.UsageDataLevel.on, jsonData.readUsageDataObjectProperty(usageDataObject.groupName, "usageDataLevel"));
    });
  });
  describe("Test writeUsageDataJsonData method", () => {
    it("should write to already existing file", () => {
      fs.writeFileSync(defaults.usageDataJsonFilePath, "");
      const usageDataLevel = usageDataObject.usageDataLevel;
      let jsonObject = {};
      jsonObject[usageDataObject.groupName] = usageDataObject.usageDataLevel;
      jsonObject = { usageDataInstances: jsonData };
      jsonObject = { usageDataInstances: { [usageDataObject.groupName]: { usageDataLevel } } };
      jsonData.writeUsageDataJsonData(usageDataObject.groupName, usageDataObject.usageDataLevel);
      assert.equal(JSON.stringify(jsonObject, null, 2), fs.readFileSync(defaults.usageDataJsonFilePath, "utf8"));
    });
  });
  describe("Test writeUsageDataJsonData method", () => {
    it("should create new existing file with correct format", () => {
      const usageDataLevel = usageDataObject.usageDataLevel;
      let jsonObject = {};
      jsonObject[usageDataObject.groupName] = usageDataObject.usageDataLevel;
      jsonObject = { usageDataInstances: jsonData };
      jsonObject = { usageDataInstances: { [usageDataObject.groupName]: { usageDataLevel } } };
      jsonData.writeUsageDataJsonData(usageDataObject.groupName, usageDataObject.usageDataLevel);
      assert.equal(JSON.stringify(jsonObject, null, 2), fs.readFileSync(defaults.usageDataJsonFilePath, "utf8"));
    });
  });
  describe("Test groupNameExists method", () => {
    it("should check if groupName exists", () => {
      fs.writeFileSync(defaults.usageDataJsonFilePath, "test");
      const usageDataLevel = usageDataObject.usageDataLevel;
      let jsonObject = {};
      jsonObject[usageDataObject.groupName] = usageDataObject.usageDataLevel;
      jsonObject = { usageDataInstances: jsonData };
      jsonObject = { usageDataInstances: { [usageDataObject.groupName]: { usageDataLevel } } };
      fs.writeFileSync(defaults.usageDataJsonFilePath, JSON.stringify((jsonObject)));
      assert.equal(true, jsonData.groupNameExists("office-addin-usage-data"));
    });
  });
  describe("Test readUsageDataSettings method", () => {
    it("should read and return parsed usage data object group settings", () => {
      const usageDataLevel = usageDataObject.usageDataLevel;
      const jsonObject = { usageDataInstances: { [usageDataObject.groupName]: { usageDataLevel } } };
      fs.writeFileSync(defaults.usageDataJsonFilePath, JSON.stringify((jsonObject)));
      assert.equal(JSON.stringify(jsonObject.usageDataInstances[usageDataObject.groupName]), JSON.stringify(jsonData.readUsageDataSettings(usageDataObject.groupName)));
    });
  });
  describe("Test reportSuccess", () => {
    it("should send success events successfully", () => {
      addInUsageData.reportSuccess("testMethod-reportSuccess", {TestVal: 42, OtherTestVal: "testing"});
      assert.equal(addInUsageData.getEventsSent(), 1);
    });
    it("should send success events successfully, even when there's no additional data", () => {
      addInUsageData.reportSuccess("testMethod-reportSuccess");
      assert.equal(addInUsageData.getEventsSent(), 1);
    });
  });
  describe("Test reportExpectedException", () => {
    it("should send successful fail events successfully", () => {
      addInUsageData.reportExpectedException("testMethod-reportExpectedException", new Error("Test"), {TestVal: 42, OtherTestVal: "testing"});
      assert.equal(addInUsageData.getEventsSent(), 1);
    });
    it("should send successful fail events successfully, even when there's no additional data", () => {
      addInUsageData.reportExpectedException("testMethod-reportExpectedException", new Error("Test"));
      assert.equal(addInUsageData.getEventsSent(), 1);
    });
  });
  describe("Test sendUsageDataEvent", () => {
    it("should send events successfully", () => {
      addInUsageData.sendUsageDataEvent({TestVal: 42, OtherTestVal: "testing"});
      assert.equal(addInUsageData.getEventsSent(), 1);
    });
    it("should send events successfully, even when there's no data", () => {
      addInUsageData.sendUsageDataEvent();
      assert.equal(addInUsageData.getEventsSent(), 1);
    });
  });
  describe("Test reportException", () => {
    it("should send exceptions successfully", () => {
      addInUsageData.reportException("testMethod-reportException", new Error("Test"), {TestVal: 42, OtherTestVal: "testing"});
      assert.equal(addInUsageData.getExceptionsSent(), 1);
    });
    it("should send exceptions successfully, even when there's no data", () => {
      addInUsageData.reportException("testMethod-reportException", new Error("Test"));
      assert.equal(addInUsageData.getExceptionsSent(), 1);
    });
  });
});
