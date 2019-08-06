import * as appInsights from "applicationinsights";
import * as assert from "assert";
import * as fs from "fs";
import * as mocha from "mocha";
import * as os from "os";
import * as path from "path";
import * as officeAddinTelemetry from "../src/officeAddinTelemetry";
import * as jsonData from "../src/telemetryJsonData";

let addInTelemetry: officeAddinTelemetry.OfficeAddinTelemetry;
const err = new Error(`this error contains a file path:C:/${os.homedir()}/AppData/Roaming/npm/node_modules//alanced-match/index.js`);
const testJsonFilePath = path.join(os.homedir(), "/mochaTest.json");
const telemetryObject: officeAddinTelemetry.ITelemetryObject = {
  groupName: "Office-Addin-Scripts",
  projectName: "Test-Project",
  instrumentationKey: "de0d9e7c-1f46-4552-bc21-4e43e489a015",
  promptQuestion: "-----------------------------------------\nDo you want to opt-in for telemetry?[y/n]\n-----------------------------------------",
  raisePrompt: false,
  telemetryLevel: officeAddinTelemetry.telemetryLevel.verbose,
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
      // delete mochaTest.json
      if (fs.existsSync(testJsonFilePath)) {
        fs.unlinkSync(testJsonFilePath);
      }
      assert.equal(jsonData.promptForTelemetry(telemetryObject.groupName, telemetryObject.telemetryJsonFilePath), true);
    });
  });

  describe("Test promptForTelemetry method", () => {
    it("Should return 'false' because testJsonFilePath exists and groupName exists in file", () => {
      jsonData.writeTelemetryJsonData(telemetryObject.groupName, telemetryObject.telemetryLevel, telemetryObject.telemetryJsonFilePath);
      assert.equal(jsonData.promptForTelemetry(telemetryObject.groupName, telemetryObject.telemetryJsonFilePath), false);
    });
  });

  describe("Test promptForTelemetry method", () => {
    it("Should return 'true' because testJsonFilePath exists but groupName doesn't exist on file", () => {
      jsonData.writeTelemetryJsonData(telemetryObject.groupName, telemetryObject.telemetryLevel, telemetryObject.telemetryJsonFilePath);
      assert.equal(jsonData.promptForTelemetry("test group name", telemetryObject.telemetryJsonFilePath), true);
    });
  });

  describe("Test telemetryOptIn method", () => {
    it("Should write out file with groupName set to true to testJsonFilePath", () => {
      addInTelemetry.telemetryOptIn(telemetryObject.testData, "y");
      const jsonTelemtryData = jsonData.readTelemetryJsonData(telemetryObject.telemetryJsonFilePath);
      assert.equal(jsonTelemtryData.telemetryInstances[telemetryObject.groupName].telemetryObjectLevel, "verbose");
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
      const test1 = { Name: { value: "test", elapsedTime: 9 } };
      addInTelemetry.reportEvent("TestData", test1, 12);
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
      jsonObject = { telemetryInstances: jsonData};
      jsonObject = { telemetryInstances: {[telemetryObject.groupName]: {telemetryLevel}} };
      fs.writeFileSync(testJsonFilePath, JSON.stringify((jsonObject)));
      const telemetryJsonData = jsonData.readTelemetryJsonData(testJsonFilePath);
      telemetryJsonData.telemetryInstances[telemetryObject.groupName]["testProperty"] = 0;
      jsonData.modifyTelemetryJsonData(telemetryObject.groupName, "testProperty", 0, testJsonFilePath);
      assert.equal( JSON.stringify(telemetryJsonData),  JSON.stringify(jsonData.readTelemetryJsonData(testJsonFilePath)));
    });
  });
  describe("Test readTelemetryJsonData method", () => {
    it("should read and return parsed object object from telemetry", () => {
      const telemetryLevel = telemetryObject.telemetryLevel;
      let jsonObject = {};
      jsonObject[telemetryObject.groupName] = telemetryObject.telemetryLevel;
      jsonObject = { telemetryInstances: jsonData};
      jsonObject = { telemetryInstances: {[telemetryObject.groupName]: {telemetryLevel}} };
      fs.writeFileSync(testJsonFilePath, JSON.stringify((jsonObject)));
      assert.equal( JSON.stringify(jsonObject),  JSON.stringify(jsonData.readTelemetryJsonData(testJsonFilePath)));
    });
  });
  describe("Test readTelemetryLevel method", () => {
    it("should read and return object's telemetry level from file", () => {
      const telemetryLevel = telemetryObject.telemetryLevel;
      let jsonObject = {};
      jsonObject[telemetryObject.groupName] = telemetryObject.telemetryLevel;
      jsonObject = { telemetryInstances: jsonData};
      jsonObject = { telemetryInstances: {[telemetryObject.groupName]: {telemetryLevel}} };
      fs.writeFileSync(testJsonFilePath, JSON.stringify((jsonObject)));
      assert.equal( officeAddinTelemetry.telemetryLevel.verbose,  jsonData.readTelemetryLevel(telemetryObject.groupName));
    });
  });
  describe("Test readTelemetryObjectProperty method", () => {
    it("should read and return parsed object object from telemetry", () => {
      const telemetryLevel = telemetryObject.telemetryLevel;
      let jsonObject = {};
      jsonObject[telemetryObject.groupName] = telemetryObject.telemetryLevel;
      jsonObject = { telemetryInstances: jsonData};
      jsonObject = { telemetryInstances: {[telemetryObject.groupName]: {telemetryLevel}} };
      fs.writeFileSync(testJsonFilePath, JSON.stringify((jsonObject)));
      assert.equal( officeAddinTelemetry.telemetryLevel.verbose,  jsonData.readTelemetryObjectProperty(telemetryObject.groupName, "telemetryObjectLevel"));
    });
  });
  describe("Test writeTelemetryJsonData method", () => {
    it("should write to already existing file", () => {
      fs.writeFileSync(testJsonFilePath, "");
      const telemetryObjectLevel = telemetryObject.telemetryLevel;
      let jsonObject = {};
      jsonObject[telemetryObject.groupName] = telemetryObject.telemetryLevel;
      jsonObject = { telemetryInstances: jsonData};
      jsonObject = { telemetryInstances: {[telemetryObject.groupName]: {telemetryObjectLevel}} };
      jsonData.writeTelemetryJsonData(telemetryObject.groupName, telemetryObjectLevel, testJsonFilePath);
      assert.equal(JSON.stringify(jsonObject, null, 2), fs.readFileSync(testJsonFilePath, "utf8"));
    });
});
describe("Test writeTelemetryJsonData method", () => {
  it("should create new existing file with correct format", () => {
    const telemetryObjectLevel = telemetryObject.telemetryLevel;
    let jsonObject = {};
    jsonObject[telemetryObject.groupName] = telemetryObject.telemetryLevel;
    jsonObject = { telemetryInstances: jsonData};
    jsonObject = { telemetryInstances: {[telemetryObject.groupName]: {telemetryObjectLevel}} };
    jsonData.writeTelemetryJsonData(telemetryObject.groupName, telemetryObject.telemetryLevel, testJsonFilePath);
    assert.equal(JSON.stringify(jsonObject, null, 2), fs.readFileSync(testJsonFilePath, "utf8"));
  });
});
describe("Test groupNameExists method", () => {
  it("should check if groupName exists", () => {
    fs.writeFileSync(testJsonFilePath, "test");
    const telemetryLevel = telemetryObject.telemetryLevel;
    let jsonObject = {};
    jsonObject[telemetryObject.groupName] = telemetryObject.telemetryLevel;
    jsonObject = { telemetryInstances: jsonData};
    jsonObject = { telemetryInstances: {[telemetryObject.groupName]: {telemetryLevel}} };
    fs.writeFileSync(testJsonFilePath, JSON.stringify((jsonObject)));
    assert.equal(true, jsonData.groupNameExists(jsonData.readTelemetryJsonData(testJsonFilePath), "Office-Addin-Scripts"));
  });
});
});
