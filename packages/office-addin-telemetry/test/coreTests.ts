import * as appInsights from "applicationinsights";
import * as assert from "assert";
import * as fs from "fs";
import * as mocha from "mocha";
import { OfficeAddinTelemetry, telemetryType } from "../src/officeAddinTelemetry";
const telemetryObject = {
  groupName: "Office-Addin-Scripts",
  instrumentationKey: "de0d9e7c-1f46-4552-bc21-4e43e489a015",
  promptQuestion: "-----------------------------------------\nDo you want to opt-in for telemetry?[y/n]\n-----------------------------------------",
  telemetryEnabled: true,
  telemetryType: telemetryType.applicationinsights,
  testData: true,
};
const addInTelemetry = new OfficeAddinTelemetry(telemetryObject);
const err = new Error("this error contains a file path:C://" + require("os").homedir() + "/AppData//Roaming//npm//node_modules//balanced-match//index.js"); // for testing parse method
const path = require("os").homedir() + "/mochaTest.json";

    describe("test reportEvent method", () => {
    it("should track event of object passed in with a project name", () => {
        addInTelemetry.setTelemetryOff();
        const test1 = { Name: { value: "julian", elapsedTime: 9 } };
        addInTelemetry.reportEvent("TestData", test1);
        assert(1 === addInTelemetry.getEventsSent());
    });
  });

    describe(" test reportError method", () => {
    it("should send telemetry execption", () => {
        addInTelemetry.setTelemetryOff();
        addInTelemetry.reportError("ReportErrorCheck",err);
        assert(1 === addInTelemetry.getExceptionsSent());
    });
  });

    describe(" test addTelemetry method", () => {
        it("should add object to telemetry", () => {
          const test = {};
          addInTelemetry.addTelemetry(test, "Name", "julian", 9);
          assert(JSON.stringify(test) === JSON.stringify({ Name: { value: "julian", elapsedTime: 9 } }));
        });
          });

    describe("test checkPrompt method", () => {
          it("should check to see if it has writen to a file if not creates file and writes to it returns true", () => {
            if(fs.existsSync(path)){
            fs.unlinkSync(path); // deletes file if it exists
          }
          assert(true === addInTelemetry.checkPrompt(path));
          });

          it("should check to see if JSON Object is in file if already created, if appropriate word(s) are not in, returns true and writes to file. Writes group name then wether telemetry is enabled or disabled", () => {
          fs.writeFileSync(path, "");
          let response;
          assert(true === addInTelemetry.checkPrompt(path));
          let jsonData = fs.readFileSync(path, "utf8");
          let projectJsonData = JSON.parse(jsonData.toString());
          if (Object.getOwnPropertyNames(projectJsonData.telemetryInstances).includes(telemetryObject.groupName) && projectJsonData.telemetryInstances[telemetryObject.groupName] === true) {
            response = true;
          } else {
            response = false;
          }

          assert(true === response);

          });

          it("should check to see if text is in file, if appropriate word(s) are in, returns false", () => {
            if (fs.existsSync(path)) {
            assert(false === addInTelemetry.checkPrompt(path));
            }
            });
        });
    describe("test telemetryOptIn method", () => {
        it("should display user asking to opt in, changes to true if user types y ", () => {
                addInTelemetry.telemetryOptIn(1);
                assert(true === addInTelemetry.telemetryOptedIn2());
          });
          it("should display user asking to opt in, changes to false if user types anything else then y ", () => {
                addInTelemetry.telemetryOptIn(2);
                assert(false === addInTelemetry.telemetryOptedIn2());
      });
          });

    describe("test setTelemetryOff method", () => {
        it("should change samplingPercentage to 100, turns telemetry on", () => {
            addInTelemetry.setTelemetryOn();
            addInTelemetry.setTelemetryOff();
            assert(0 === appInsights.defaultClient.config.samplingPercentage);
        });
          });
    describe("test setTelemetryOn method", () => {
        it("should change samplingPercentage to 100, turns telemetry on", () => {
            addInTelemetry.setTelemetryOff();
            addInTelemetry.setTelemetryOn();
          assert(100 === appInsights.defaultClient.config.samplingPercentage);
        });
          });
    describe("test isTelemetryOn method", () => {
        it("should return true if samplingPercentage is on(100)", () => {
              appInsights.defaultClient.config.samplingPercentage = 100;
            assert(true === addInTelemetry.isTelemetryOn());
          });
      it("should return false if samplingPercentage is off(0)", () => {
              appInsights.defaultClient.config.samplingPercentage = 0;
            assert(false === addInTelemetry.isTelemetryOn());
        });
        });
    describe("test getTelemetryKey method", () => {
      it("should return telemetry key", () => {
            assert("de0d9e7c-1f46-4552-bc21-4e43e489a015" === addInTelemetry.getTelemetryKey());
        });
          });

    describe("test getEventsSent method", () => {
        it("should return amount of events successfully sent", () => {
                addInTelemetry.setTelemetryOff();
                const test1 = { Name: { value: "julian", elapsedTime: 9 } };
                addInTelemetry.reportEvent("TestData", test1, 12);
                assert(1 === addInTelemetry.getEventsSent());
        });
          });

    describe("test getExceptionsSent method", () => {
        it("should return amount of exceptions successfully sent ",() => {
                addInTelemetry.setTelemetryOff();
                addInTelemetry.reportError("TestData", err);
                assert(2 === addInTelemetry.getExceptionsSent());
            });
          });

    describe("test telemetryOptedIn method", () => {
        it("should return true if user opted in", () => {
            addInTelemetry.telemetryOptIn(1);
            assert(true === addInTelemetry.telemetryOptedIn2());
          });
          it("should return false if user opted out", () => {
            addInTelemetry.telemetryOptIn(2);
            assert(false === addInTelemetry.telemetryOptedIn2());
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
    at Object.<anonymous> (coreTests.ts:15:13)`;
    addInTelemetry.testMaskFilePaths(err);
                  assert(compareError.name ===  err.name);
                  assert(compareError.message ===  err.message);
                  assert(err.stack.includes(compareError.stack));
              });
            });
