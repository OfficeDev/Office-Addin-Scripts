import { OfficeAddinTelemetry, telemetryType,telemetryObject } from "./officeAddinTelemetry";
import * as os from 'os';
import * as fs from "fs";
import * as appInsights from "applicationinsights";

import { Base, ExceptionData, ExceptionDetails, StackFrame } from 'applicationinsights/out/Declarations/Contracts';
import { inherits } from 'util';
import * as readline from "readline";//used
import { getMaxListeners } from "cluster";
const telemetryObject1 = {
    instrumentationKey: "de0d9e7c-1f46-4552-bc21-4e43e489a015",
    telemetryType: telemetryType.applicationinsights,
    groupName: "Office-Addin-Scripts",
    promptQuestion: "-----------------------------------------\nDo you want to opt-in for telemetry?[y/n]\n-----------------------------------------",
    telemetryEnabled: true,
    testData: false
  }

  const telemetryObject2 = {
    instrumentationKey: "de0d9e7c-1f46-4552-bc21-4e43e489a015",
    telemetryType: telemetryType.applicationinsights,
    groupName: "Office-Addin-Scriptz",
    promptQuestion: "-----------------------------------------\nTester\n-----------------------------------------",
    telemetryEnabled: true,
    testData: false
  }
  const telemetryObject3 = {
    instrumentationKey: "test",
    telemetryType: telemetryType.applicationinsights,
    groupName: "huh",
    promptQuestion: "Yikes",
    telemetryEnabled: true,
    testData: false
};
  const telemetryObject4 = {
    instrumentationKey: "test2",
    telemetryType: telemetryType.applicationinsights,
    groupName: "huh2",
    promptQuestion: "Yikes2",
    telemetryEnabled: true,
    testData: false
  }
  const telemetryObject5 = {
    instrumentationKey: "de0d9e7c-1f46-4552-bc21-4e43e489a015",
    telemetryType: telemetryType.applicationinsights,
    groupName: "Office-Addin-TooloBox",
    promptQuestion: "-----------------------------------------\nDo you want to Opt-in?\n-----------------------------------------",
    telemetryEnabled: true,
    testData: false
  }
  /*fs.writeFileSync(require('os').homedir() + "/tester.txt", JSON.stringify({ JSON.parsGroupName1: telemetryObject1 }, null, 2));
  var text = fs.readFileSync(require('os').homedir() + "/tester.txt", "utf8");
  fs.writeFileSync(require('os').homedir() + "/tester.txt", text + JSON.stringify({ "GroupName2": telemetryObject2 }, null, 2));
  text = fs.readFileSync(require('os').homedir() + "/tester.txt", "utf8");
  var old = text.substring(text.indexOf("GroupName1", 0), text.indexOf("}", text.indexOf("GroupName2"))+ 1);
  var updated = text.replace('"telemetryEnabled": true,', '"telemetryEnabled": false,');
  fs.writeFileSync(require('os').homedir() + "/tester.txt", text.replace(old,updated));*/
const test = new OfficeAddinTelemetry(telemetryObject1);
const addInTelemetry1 = new OfficeAddinTelemetry(telemetryObject2);
const addInTelemetry2 = new OfficeAddinTelemetry(telemetryObject3,"Office-Addin-Scriptz");
const addInTelemetry3 = new OfficeAddinTelemetry(telemetryObject4, "Office-Addin-Scripts");
const addInTelemetry4 = new OfficeAddinTelemetry(telemetryObject5);

/*const testData = {
    "Host": "Excel",
    "IsTestData":true,
    "IsTypeScript": true,
    "OperatingSystem":os.platform(),
};
console.log(Object.entries(testData));
var test = {};
addInTelemetry.addTelemetry(test, "Name", "julian", 9);
//console.log(test);
addInTelemetry.addTelemetry(test, "Host", "Microsoft", 10);
addInTelemetry.addTelemetry(test, "Script", "typescript", 11);
//addInTelemetry.reportEvent("TestData",test);
//addInTelemetry.addTelemetry(tester,"Typescript", true);
const exception = new Error("this error contains a file path: (C://Users//t-juflor//AppData//Roaming//npm//node_modules//balanced-match//index.js)");
//addInTelemetry.reportError("TestData,remove sensitive info",exception);
//addInTelemetry.reportEvent("TestData,remove sensitive info",test1);
addInTelemetry.setTelemetryOff();
//console.log(addInTelemetry.isTelemetryOn());
//addInTelemetry.reportEvent("Tester", test1);
addInTelemetry.setTelemetryOn();
//console.log(addInTelemetry.isTelemetryOn());
//addInTelemetry.sendTelemetryEvents("Should make it!", test1);
addInTelemetry.setTelemetryOff();
//console.log(addInTelemetry.isTelemetryOn());
//addInTelemetry.sendTelemetryEvents("Shouldn't make it", test2);
addInTelemetry.setTelemetryOn();
//console.log(addInTelemetry.isTelemetryOn());
//addInTelemetry.sendTelemetryEvents("Should make it part2!", test2);
*/