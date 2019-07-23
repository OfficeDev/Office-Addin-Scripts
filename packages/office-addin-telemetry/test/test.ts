import * as appInsights from "applicationinsights";
import * as assert from "assert";
import * as fs from "fs";
import * as mocha from "mocha";
import * as os from "os";
import * as path from "path";
import * as jsonData from "../src/jsonData";
import * as officeAddinTelemetry from "../src/officeAddinTelemetry";

let addInTelemetry: officeAddinTelemetry.OfficeAddinTelemetry;
const err = new Error(`this error contains a file path:C:/${os.homedir()}/AppData/Roaming/npm/node_modules//alanced-match/index.js`);
const testJsonFilePath = path.join(os.homedir(), "/mochaTest.json");
const telemetryObject: officeAddinTelemetry.ITelemetryObject = {
  groupName: "Office-Addin-Scripts",
  instrumentationKey: "de0d9e7c-1f46-4552-bc21-4e43e489a015",
  promptQuestion: "-----------------------------------------\nDo you want to opt-in for telemetry?[y/n]\n-----------------------------------------",
  raisePrompt: false,
  telemetryEnabled: true,
  telemetryJsonFilePath: testJsonFilePath,
  telemetryType: officeAddinTelemetry.telemetryType.applicationinsights,
  testData: true,
};

describe("Test office-addin-telemetry-package", function() {
  beforeEach(function() {
    if (fs.existsSync(testJsonFilePath)) {
      fs.unlinkSync(testJsonFilePath);
    }
    addInTelemetry = new officeAddinTelemetry.OfficeAddinTelemetry(telemetryObject);
  });
  describe("Test reportEvent method", () => {
    it("should track event of object passed in with a project name", () => {
      const testEvent = { Name: { value: "testValue", elapsedTime: 9 } };
      addInTelemetry.reportEvent("testData", testEvent);
      assert.equal(addInTelemetry.getEventsSent(), 1);
    });
  });

  describe("Test reportError method", () => {
    it("should send telemetry exception", () => {
      addInTelemetry.reportError("ReportErrorCheck", err);
      assert.equal(addInTelemetry.getExceptionsSent(), 1);
    });
  });

  describe("Test addTelemetry method", () => {
    it("should add object to telemetry", () => {
      const testObject = {};
      addInTelemetry.addTelemetry(testObject, "Name", "testValue", 9);
      assert.equal(JSON.stringify(testObject), JSON.stringify({ Name: { value: "testValue", elapsedTime: 9 } }));
    });
  });

  describe("Test deleteTelemetry method", () => {
    it("should delete object from telemetry", () => {
      const testObject = {};
      addInTelemetry.addTelemetry(testObject, "Name", "testvalue", 9);
      addInTelemetry.deleteTelemetry(testObject, "Name");
      assert.equal(JSON.stringify(testObject), JSON.stringify({}));
    });
  });

  describe("Test promptForTelemetry method", () => {
    it("Should return 'true' because testJsonFilePath doesn't exist", () => {
      // Delete mochaTest.json
      if (fs.existsSync(testJsonFilePath)) {
        fs.unlinkSync(testJsonFilePath);
      }
      assert.equal(jsonData.promptForTelemetry(telemetryObject.groupName, telemetryObject.telemetryJsonFilePath), true);
    });
  });

  describe("Test promptForTelemetry method", () => {
    it("Should return 'false' because testJsonFilePath exists and groupName exists in file", () => {
      jsonData.writeNewTelemetryJsonFile(telemetryObject.groupName, telemetryObject.telemetryEnabled, telemetryObject.telemetryJsonFilePath);
      assert.equal(jsonData.promptForTelemetry(telemetryObject.groupName, telemetryObject.telemetryJsonFilePath), false);
    });
  });

  describe("Test promptForTelemetry method", () => {
    it("Should return 'false' because testJsonFilePath exists and groupName doesn't exists in file", () => {
      jsonData.writeNewTelemetryJsonFile("testGroupName", telemetryObject.telemetryEnabled, telemetryObject.telemetryJsonFilePath);
      assert.equal(jsonData.promptForTelemetry(telemetryObject.groupName, telemetryObject.telemetryJsonFilePath), true);
    });
  });

  describe("Test telemetryOptIn method", () => {
    it("Should write out file with groupName set to true to testJsonFilePath", () => {
      addInTelemetry.telemetryOptIn(telemetryObject.testData, "y");
      const jsonTelemtryData = jsonData.readTelemetryJsonData(telemetryObject.telemetryJsonFilePath);
      assert.equal(jsonTelemtryData.telemetryInstances[telemetryObject.groupName], true);
    });
  });

  describe("test setTelemetryOff method", () => {
    it("should change samplingPercentage to 100, turns telemetry on", () => {
      addInTelemetry.setTelemetryOn();
      addInTelemetry.setTelemetryOff();
      assert.equal(addInTelemetry.isTelemetryOn(), false);
    });
  });

  describe("test setTelemetryOn method", () => {
    it("should change samplingPercentage to 100, turns telemetry on", () => {
      addInTelemetry.setTelemetryOff();
      addInTelemetry.setTelemetryOn();
      assert.equal(appInsights.defaultClient.config.samplingPercentage, 100);
    });
  });

  describe("test isTelemetryOn method", () => {
    it("should return true if samplingPercentage is on(100)", () => {
      appInsights.defaultClient.config.samplingPercentage = 100;
      assert.equal(addInTelemetry.isTelemetryOn(), true);
    });
    it("should return false if samplingPercentage is off(0)", () => {
      appInsights.defaultClient.config.samplingPercentage = 0;
      assert.equal(addInTelemetry.isTelemetryOn(), false);
    });
  });

  describe("test getTelemetryKey method", () => {
    it("should return telemetry key", () => {
      assert.equal(addInTelemetry.getTelemetryKey(), "de0d9e7c-1f46-4552-bc21-4e43e489a015");
    });
  });

  describe("test getEventsSent method", () => {
    it("should return amount of events successfully sent", () => {
      addInTelemetry.setTelemetryOff();
      const test1 = { Name: { value: "test", elapsedTime: 9 } };
      addInTelemetry.reportEvent("TestData", test1, 12);
      assert.equal(addInTelemetry.getEventsSent(), 1);
    });
  });

  describe("test getExceptionsSent method", () => {
    it("should return amount of exceptions successfully sent ", () => {
      addInTelemetry.setTelemetryOff();
      addInTelemetry.reportError("TestData", err);
      assert.equal(addInTelemetry.getExceptionsSent(), 1);
    });
  });

  describe("test parseErrors method", () => {
    it("should return a parsed file path error", () => {
      addInTelemetry.setTelemetryOff();
      const compareError = new Error();
      compareError.name = "TestData";
      compareError.message = "this error contains a file path:C:index.js";
      // may throw error if change any part of the top of the test file
      compareError.stack = `ReportErrorCheck: this error contains a file path:C:index.js
    at Object.<anonymous> (test.ts:11:13)`;
      addInTelemetry.maskFilePaths(err);
      assert.equal(compareError.name, err.name);
      assert.equal(compareError.message, err.message);
      assert.equal(err.stack.includes(compareError.stack), true);
    });
  });
});
