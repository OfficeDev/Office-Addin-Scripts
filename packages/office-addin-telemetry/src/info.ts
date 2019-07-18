import { OfficeAddinTelemetry, telemetryType,telemetryObject,promptForTelemetry } from "./officeAddinTelemetry";
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
    promptQuestion: "-----------------------------------------\nWe collecte diaganoistic data but would welcome the collection of custom events. Would you like to opt-in for the extended telemetry?[y/n]\n-----------------------------------------",
    telemetryEnabled: false,
    testData: false,
  }
promptForTelemetry("Office-Addin-Scripts", false);
var tester = {};

const test = new OfficeAddinTelemetry(telemetryObject1);
test.addTelemetry(tester, "delete", "shouldn't be here");
test.addTelemetry(tester, "keep", "should be here");
test.deleteTelemetry(tester, "delete");
console.log(JSON.stringify(tester));
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
*/