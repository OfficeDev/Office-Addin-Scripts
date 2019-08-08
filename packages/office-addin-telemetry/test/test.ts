import * as appInsights from "applicationinsights";
import * as assert from "assert";
import * as fs from "fs";
import * as mocha from "mocha";
import * as os from "os";
import * as defaults from "../src/defaults";
import * as officeAddinTelemetry from "../src/officeAddinTelemetry";
import * as jsonData from "../src/telemetryJsonData";

let addInTelemetry: officeAddinTelemetry.OfficeAddinTelemetry;
const err = new Error(`this error contains a file path:C:/${os.homedir()}/AppData/Roaming/npm/node_modules//alanced-match/index.js`);
let telemetryData: string;
const telemetryObject: officeAddinTelemetry.ITelemetryOptions = {
  groupName: "Office-Addin-Telemetry",
  projectName: "Test-Project",
  instrumentationKey: "de0d9e7c-1f46-4552-bc21-4e43e489a015",
  promptQuestion: "-----------------------------------------\nDo you want to opt-in for telemetry?[y/n]\n-----------------------------------------",
  raisePrompt: false,
  telemetryLevel: officeAddinTelemetry.telemetryLevel.verbose,
  telemetryType: officeAddinTelemetry.telemetryType.applicationinsights,
  testData: true,
};

describe("Test office-addin-telemetry-package", function() {
  this.beforeAll(function() {
    try {
    if (fs.existsSync(defaults.telemetryJsonFilePath) && fs.readFileSync(defaults.telemetryJsonFilePath, "utf8").toString() !== undefined) {
      telemetryData = JSON.parse(fs.readFileSync(defaults.telemetryJsonFilePath, "utf8").toString());
    }
  } catch {
    telemetryData = undefined;
  }
  });
  this.afterAll(function() {
    if (fs.existsSync(defaults.telemetryJsonFilePath)) {
      fs.writeFileSync(defaults.telemetryJsonFilePath, JSON.stringify((telemetryData), null, 2));
    }
  });
  beforeEach(function() {
    if (fs.existsSync(defaults.telemetryJsonFilePath)) {
      fs.unlinkSync(defaults.telemetryJsonFilePath);
    }
    addInTelemetry = new officeAddinTelemetry.OfficeAddinTelemetry(telemetryObject);
  });
  describe("Test reportEvent method", () => {
    it("should track event of object passed in with a project name", () => {
      const testEvent = {
        Test1: [true, 100],
        ScriptType: ["Java", 1],
      };
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

  describe("Test promptForTelemetry method", () => {
    it("Should return 'true' because telemetryJsonFilePath doesn't exist", () => {
      // delete officeAddinTelemetry.json
      if (fs.existsSync(defaults.telemetryJsonFilePath)) {
        fs.unlinkSync(defaults.telemetryJsonFilePath);
      }
      assert.equal(jsonData.needToPromptForTelemetry(telemetryObject.groupName), true);
    });
  });

  describe("Test promptForTelemetry method", () => {
    it("Should return 'false' because telemetryJsonFilePath exists and groupName exists in file", () => {
      jsonData.writeTelemetryJsonData(telemetryObject.groupName, telemetryObject.telemetryLevel);
      assert.equal(jsonData.needToPromptForTelemetry(telemetryObject.groupName), false);
    });
  });

  describe("Test promptForTelemetry method", () => {
    it("Should return 'true' because telemetryJsonFilePath exists but groupName doesn't exist on file", () => {
      jsonData.writeTelemetryJsonData(telemetryObject.groupName, telemetryObject.telemetryLevel);
      assert.equal(jsonData.needToPromptForTelemetry("test group name"), true);
    });
  });

  describe("Test telemetryOptIn method", () => {
    it("Should write out file with groupName set to true to telemetryJsonFilePath", () => {
      addInTelemetry.telemetryOptIn(telemetryObject.testData, "y");
      const jsonTelemtryData = jsonData.readTelemetryJsonData();
      assert.equal(jsonTelemtryData.telemetryInstances[telemetryObject.groupName].telemetryLevel, telemetryObject.telemetryLevel);
    });
  });

  describe("Test setTelemetryOff method", () => {
    it("should change samplingPercentage to 100, turns telemetry on", () => {
      addInTelemetry.setTelemetryOn();
      addInTelemetry.setTelemetryOff();
      assert.equal(addInTelemetry.isTelemetryOn(), false);
    });
  });

  describe("Test setTelemetryOn method", () => {
    it("should change samplingPercentage to 100, turns telemetry on", () => {
      addInTelemetry.setTelemetryOff();
      addInTelemetry.setTelemetryOn();
      assert.equal(appInsights.defaultClient.config.samplingPercentage, 100);
    });
  });

  describe("Test isTelemetryOn method", () => {
    it("should return true if samplingPercentage is on(100)", () => {
      appInsights.defaultClient.config.samplingPercentage = 100;
      assert.equal(addInTelemetry.isTelemetryOn(), true);
    });
    it("should return false if samplingPercentage is off(0)", () => {
      appInsights.defaultClient.config.samplingPercentage = 0;
      assert.equal(addInTelemetry.isTelemetryOn(), false);
    });
  });

  describe("Test getTelemetryKey method", () => {
    it("should return telemetry key", () => {
      assert.equal(addInTelemetry.getTelemetryKey(), "de0d9e7c-1f46-4552-bc21-4e43e489a015");
    });
  });

  describe("Test getEventsSent method", () => {
    it("should return amount of events successfully sent", () => {
      addInTelemetry.setTelemetryOff();
      const testEvent = {
        Test1: [true, 100],
        ScriptType: ["Java", 1],
      };
      addInTelemetry.reportEvent("TestData", testEvent);
      assert.equal(addInTelemetry.getEventsSent(), 1);
    });
  });

  describe("Test getExceptionsSent method", () => {
    it("should return amount of exceptions successfully sent ", () => {
      addInTelemetry.setTelemetryOff();
      addInTelemetry.reportError("TestData", err);
      assert.equal(addInTelemetry.getExceptionsSent(), 1);
    });
  });

  describe("Test telemetryLevel method", () => {
    it("should return the telemetry level of the object", () => {
      assert.equal("verbose", addInTelemetry.telemetryLevel());
    });
  });
  describe("Test maskFilePaths method", () => {
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
  describe("Test modifySetting method", () => {
    it("should modify or create specific property to new value", () => {
      const telemetryLevel = telemetryObject.telemetryLevel;
      let jsonObject = {};
      jsonObject[telemetryObject.groupName] = telemetryObject.telemetryLevel;
      jsonObject = { telemetryInstances: jsonData };
      jsonObject = { telemetryInstances: { [telemetryObject.groupName]: { telemetryLevel } } };
      fs.writeFileSync(defaults.telemetryJsonFilePath, JSON.stringify((jsonObject)));
      const telemetryJsonData = jsonData.readTelemetryJsonData();
      const testPropertyName = "testProperty";
      telemetryJsonData.telemetryInstances[telemetryObject.groupName][testPropertyName] = 0;
      jsonData.modifyTelemetryJsonData(telemetryObject.groupName, testPropertyName, 0);
      assert.equal(JSON.stringify(telemetryJsonData), JSON.stringify(jsonData.readTelemetryJsonData()));
    });
  });
  describe("Test readTelemetryJsonData method", () => {
    it("should read and return parsed object object from telemetry", () => {
      const telemetryLevel = telemetryObject.telemetryLevel;
      let jsonObject = {};
      jsonObject[telemetryObject.groupName] = telemetryObject.telemetryLevel;
      jsonObject = { telemetryInstances: jsonData };
      jsonObject = { telemetryInstances: { [telemetryObject.groupName]: { telemetryLevel } } };
      fs.writeFileSync(defaults.telemetryJsonFilePath, JSON.stringify((jsonObject)));
      assert.equal(JSON.stringify(jsonObject), JSON.stringify(jsonData.readTelemetryJsonData()));
    });
  });
  describe("Test readTelemetryLevel method", () => {
    it("should read and return object's telemetry level from file", () => {
      const telemetryLevel = telemetryObject.telemetryLevel;
      let jsonObject = {};
      jsonObject[telemetryObject.groupName] = telemetryObject.telemetryLevel;
      jsonObject = { telemetryInstances: jsonData };
      jsonObject = { telemetryInstances: { [telemetryObject.groupName]: { telemetryLevel } } };
      fs.writeFileSync(defaults.telemetryJsonFilePath, JSON.stringify((jsonObject)));
      assert.equal(officeAddinTelemetry.telemetryLevel.verbose, jsonData.readTelemetryLevel(telemetryObject.groupName));
    });
  });
  describe("Test readTelemetryObjectProperty method", () => {
    it("should read and return parsed object object from telemetry", () => {
      const telemetryLevel = telemetryObject.telemetryLevel;
      let jsonObject = {};
      jsonObject[telemetryObject.groupName] = telemetryObject.telemetryLevel;
      jsonObject = { telemetryInstances: jsonData };
      jsonObject = { telemetryInstances: { [telemetryObject.groupName]: { telemetryLevel } } };
      fs.writeFileSync(defaults.telemetryJsonFilePath, JSON.stringify((jsonObject)));
      assert.equal(officeAddinTelemetry.telemetryLevel.verbose, jsonData.readTelemetryObjectProperty(telemetryObject.groupName, "telemetryLevel"));
    });
  });
  describe("Test writeTelemetryJsonData method", () => {
    it("should write to already existing file", () => {
      fs.writeFileSync(defaults.telemetryJsonFilePath, "");
      const telemetryLevel = telemetryObject.telemetryLevel;
      let jsonObject = {};
      jsonObject[telemetryObject.groupName] = telemetryObject.telemetryLevel;
      jsonObject = { telemetryInstances: jsonData };
      jsonObject = { telemetryInstances: { [telemetryObject.groupName]: { telemetryLevel } } };
      jsonData.writeTelemetryJsonData(telemetryObject.groupName, telemetryObject.telemetryLevel);
      assert.equal(JSON.stringify(jsonObject, null, 2), fs.readFileSync(defaults.telemetryJsonFilePath, "utf8"));
    });
  });
  describe("Test writeTelemetryJsonData method", () => {
    it("should create new existing file with correct format", () => {
      const telemetryLevel = telemetryObject.telemetryLevel;
      let jsonObject = {};
      jsonObject[telemetryObject.groupName] = telemetryObject.telemetryLevel;
      jsonObject = { telemetryInstances: jsonData };
      jsonObject = { telemetryInstances: { [telemetryObject.groupName]: { telemetryLevel } } };
      jsonData.writeTelemetryJsonData(telemetryObject.groupName, telemetryObject.telemetryLevel);
      assert.equal(JSON.stringify(jsonObject, null, 2), fs.readFileSync(defaults.telemetryJsonFilePath, "utf8"));
    });
  });
  describe("Test groupNameExists method", () => {
    it("should check if groupName exists", () => {
      fs.writeFileSync(defaults.telemetryJsonFilePath, "test");
      const telemetryLevel = telemetryObject.telemetryLevel;
      let jsonObject = {};
      jsonObject[telemetryObject.groupName] = telemetryObject.telemetryLevel;
      jsonObject = { telemetryInstances: jsonData };
      jsonObject = { telemetryInstances: { [telemetryObject.groupName]: { telemetryLevel } } };
      fs.writeFileSync(defaults.telemetryJsonFilePath, JSON.stringify((jsonObject)));
      assert.equal(true, jsonData.groupNameExists("Office-Addin-Telemetry"));
    });
  });
  describe("Test readTelemetrySettings method", () => {
    it("should read and return parsed telemetry object group settings", () => {
      const telemetryLevel = telemetryObject.telemetryLevel;
      const jsonObject = { telemetryInstances: { [telemetryObject.groupName]: { telemetryLevel } } };
      fs.writeFileSync(defaults.telemetryJsonFilePath, JSON.stringify((jsonObject)));
      assert.equal(JSON.stringify(jsonObject.telemetryInstances[telemetryObject.groupName]), JSON.stringify(jsonData.readTelemetrySettings(telemetryObject.groupName)));
    });
  });

});
