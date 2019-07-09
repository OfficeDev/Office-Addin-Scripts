import { OfficeAddinTelemetry } from "./officeAddinTelemetry";
import * as os from 'os';
const addInTelemetry = new OfficeAddinTelemetry("de0d9e7c-1f46-4552-bc21-4e43e489a015");
/*const testData = {
    "Host": "Excel",
    "IsTestData":true,
    "IsTypeScript": true,
    "OperatingSystem":os.platform(),
};
console.log(Object.entries(testData));*/
var test1 ={};
var test2 = {};
addInTelemetry.addTelemetry(test1, "OperatingSystem", os.platform());
addInTelemetry.addTelemetry(test2, "Tester2", true);
//addInTelemetry.addTelemetry(tester,"Typescript", true);
const exception = new Error("this error contains a file path: C://Users//t-juflor//AppData//Roaming//npm//node_modules//balanced-match//index.js");

addInTelemetry.reportEvent("TestData,remove sensitive info",test1);
addInTelemetry.setTelemetryOff();
//console.log(addInTelemetry.isTelemetryOn());
addInTelemetry.reportEvent("Tester", test1);
addInTelemetry.setTelemetryOn();
//console.log(addInTelemetry.isTelemetryOn());
//addInTelemetry.sendTelemetryEvents("Should make it!", test1);
addInTelemetry.setTelemetryOff();
//console.log(addInTelemetry.isTelemetryOn());
//addInTelemetry.sendTelemetryEvents("Shouldn't make it", test2);
addInTelemetry.setTelemetryOn();
//console.log(addInTelemetry.isTelemetryOn());
//addInTelemetry.sendTelemetryEvents("Should make it part2!", test2);
