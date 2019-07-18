import * as appInsights from "applicationinsights";
import { Base, ExceptionData, ExceptionDetails, StackFrame } from "applicationinsights/out/Declarations/Contracts";
import * as fs from "fs";
import * as telemetryJsonData from "./telemetryJsonData";
export enum telemetryType {
  applicationinsights = "applicationInsights",
  OtelJs = "OtelJs",
}
export interface telemetryObject {
  groupName: string;
  instrumentationKey: string;
  promptQuestion: string;
  telemetryEnabled: boolean;
  telemetryType: telemetryType;
  testData: boolean;
}
// telemetryJsonData.
export function promptForTelemetry(groupName: string, enabledTelemetry: boolean, newFileLocation = require("os").homedir() + "/officeAddinTelemetry.json") {
const chalk = require("chalk");
try {
  this.path = newFileLocation;
  const name = this.telemetryObject.groupName;
  if (fs.existsSync(this.path)) {
    const jsonData = fs.readFileSync(this.path, "utf8");
    if (jsonData !== "") {
    const projectJsonData = JSON.parse(jsonData.toString());
    if (Object.getOwnPropertyNames(projectJsonData.telemetryInstances).includes(this.telemetryObject.groupName)) {
      if (projectJsonData.telemetryInstances[name] === false) {
        this.telemetryObject.telemetryEnabled  = false;
      } else {
        this.telemetryObject.telemetryEnabled = true;
      }
    } else {
      projectJsonData.telemetryInstances[name] =  this.telemetryObject.telemetryEnabled;
      fs.writeFileSync(this.path, JSON.stringify((projectJsonData), null, 2));
    }
  } else {
    const projectJsonData = {};
    projectJsonData[name] =  this.telemetryObject.telemetryEnabled;
    fs.writeFileSync(this.path, JSON.stringify(({ "telemetryInstances" : projectJsonData}), null, 2));
  }
} else {
    const projectJsonData = {};
    projectJsonData[name] =  this.telemetryObject.telemetryEnabled;
    fs.writeFileSync(this.path, JSON.stringify(({"telemetryInstances": projectJsonData}), null, 2));
  }

} catch (err) {
  console.log(chalk.red(err));
}

}

export class OfficeAddinTelemetry {
  public  chalk = require("chalk");
  private telemetryClient = appInsights.defaultClient;
  private telemetrySource = "";
  private eventsSent = 0;
  private exceptionsSent = 0;
  private path = require("os").homedir() + "/officeAddinTelemetry.json";
  private telemetryObject;

  constructor(telemetryObject1: telemetryObject) {
    // checks to make sure it only displays the opt-in message once
    try{
    this.telemetryObject =  {
      groupName: telemetryObject1.groupName,
      instrumentationKey: telemetryObject1.instrumentationKey,
      promptQuestion: telemetryObject1.promptQuestion,
      telemetryEnabled: telemetryObject1.telemetryEnabled,
      telemetryType: telemetryObject1.telemetryType,
      testData: telemetryObject1.testData,
    }
  } catch (err){
    console.log(this.chalk.red("You did not enter all the correct information!"));
    console.log(err);
  }
    // declaring telemetry structure
    if (this.telemetryObject.telemetryType === telemetryType.applicationinsights) {
      this.telemetrySource = telemetryType.applicationinsights;
    } else if (this.telemetryObject.telmetryType.toLowerCase() === "oteljs") {
      this.telemetrySource = telemetryType.OtelJs;
    }
    if (!this.telemetryObject.testData && this.checkPrompt()) {
      this.telemetryOptIn();
    }
    this.path = require("os").homedir() + "/officeAddinTelemetry.json";
    appInsights.setup(this.telemetryObject.instrumentationKey)
      .setAutoCollectConsole(true)
      .setAutoCollectExceptions(false)
      .start()
    this.telemetryClient = appInsights.defaultClient;
    this.removeSensitiveInformation();
  }

  public async reportEvent(eventName: string, data: object, timeElapsed = 0): Promise<void> {
    if (this.telemetryOptedIn()) {
      if (this.telemetrySource === telemetryType.applicationinsights) {
        this.reportEventApplicationInsights(eventName, data);
      } else if (this.telemetrySource === telemetryType.OtelJs) {
        console.log(this.chalk.red("Feature has not been yet created."));
      }
    }
  }

  public async reportEventApplicationInsights(eventName: string, data: object): Promise<void> {
    if (this.telemetryOptedIn()) {
      for (let [key, {value, elapsedTime}] of Object.entries(data)) {
        try {
          if (!this.telemetryObject.testData) {
            this.telemetryClient.trackEvent({ name: eventName, properties: { [key]: value}, measurements: { DurationElapsed: elapsedTime}});
          }
          this.eventsSent++;
        } catch (err) {
          this.reportError("sendTelemetryEvents", err);
        }
      }
    }
  }

  public async reportError(errorName: string, err: Error): Promise<void> {
        if (this.telemetrySource === telemetryType.applicationinsights) {
          this.reportErrorApplicationInsights(errorName, err);
        } else if (this.telemetrySource === telemetryType.OtelJs) {
        console.log(this.chalk.red("Feature has not been yet created."));
        }
  }

  public async reportErrorApplicationInsights(eventName: string, err: Error): Promise<void> {
      err.name = eventName;
      if (this.telemetryObject.testData) {
        err.name = eventName;
      }
      this.telemetryClient.trackException({ exception: this.maskFilePaths(err) });
      this.exceptionsSent++;
  }

  public addTelemetry(data: { [k: string]: any}, key: string, value: any, elapsedTime: any = 0): object {
    data[key] = {value, elapsedTime};
    return data;
  }
  public deleteTelemetry(data: { [k: string]: any}, key: string, value: any): object {
    delete data[key];
    return data;
  }

  public checkPrompt(newFileLocation = this.path): boolean {
    try {
      this.path = newFileLocation;
      const name = this.telemetryObject.groupName;
      if (fs.existsSync(this.path)) {
        const jsonData = fs.readFileSync(this.path, "utf8");
        if (jsonData !== ""){
        const projectJsonData = JSON.parse(jsonData.toString());
        if (Object.getOwnPropertyNames(projectJsonData.telemetryInstances).includes(this.telemetryObject.groupName)) {
          if (projectJsonData.telemetryInstances[name] === false) {
            this.telemetryObject.telemetryEnabled  = false;
          }else{
            this.telemetryObject.telemetryEnabled = true;
          }
          return false;
        } else {
          projectJsonData.telemetryInstances[name] =  this.telemetryObject.telemetryEnabled;
          fs.writeFileSync(this.path, JSON.stringify((projectJsonData), null, 2));
          return true;
        }
      } else {
        const projectJsonData = {};
        projectJsonData[name] =  this.telemetryObject.telemetryEnabled;
        fs.writeFileSync(this.path, JSON.stringify(({"telemetryInstances": projectJsonData}), null, 2));
      }
    } else {
        const projectJsonData = {};
        projectJsonData[name] =  this.telemetryObject.telemetryEnabled;
        fs.writeFileSync(this.path, JSON.stringify(({"telemetryInstances": projectJsonData}), null, 2));
      }
    } catch (err) {
      this.reportError("checkPrompt", err);
    }

    return true;

  }

  public telemetryOptIn(testValue = 0): void { //FIX TEST
    const readlineSync = require("readline-sync");
    const chalk = require("chalk");
    if (this.telemetryObject.promptQuestion === undefined || this.telemetryObject.promptQuestion === "") {
      this.telemetryObject.telemetryEnabled = false;
    } else {
      const name = this.telemetryObject.groupName;
      const jsonData = fs.readFileSync(this.path, "utf8");
      const projectJsonData = JSON.parse(jsonData.toString());
    let response = "";
    if (testValue === 0) {
      try {
        response = readlineSync.question(chalk.blue(this.telemetryObject.promptQuestion));
      } catch (err) {
        this.reportError("TelemetryOptIn", err);
      }
    }
    if (response.toLowerCase() === "y" || testValue === 1) {
      this.telemetryObject.telemetryEnabled = true;
      projectJsonData.telemetryInstances[name] = true;
      fs.writeFileSync(this.path, JSON.stringify((projectJsonData), null, 2));
      console.log(chalk.green("Telemetry will be sent!"));
    } else {
      this.telemetryObject.telemetryEnabled = false;
      projectJsonData.telemetryInstances[name] = false;
      fs.writeFileSync(this.path, JSON.stringify((projectJsonData), null, 2));
      console.log(chalk.red("You will not be sending telemetry"));
    }
  }
  }
  public setTelemetryOff() {
    appInsights.defaultClient.config.samplingPercentage = 0;
  }
  public setTelemetryOn() {
    appInsights.defaultClient.config.samplingPercentage = 100;
  }
  public isTelemetryOn(): boolean {
    if (appInsights.defaultClient.config.samplingPercentage === 100) {
      return true;
    } else {
      return false;
    }
  }
  public getTelemetryKey(): string {
    return this.telemetryObject.instrumentationKey;
  }
  public getTelemetryObject(): Object {
    return this.telemetryObject;
  }
  public getEventsSent(): any {
    return this.eventsSent;
  }
  public getExceptionsSent(): any {
    return this.exceptionsSent;
  }

  public telemetryOptedIn2(): boolean {// for mocha test
    return this.telemetryObject.telemetryEnabled;
  }
  public testMaskFilePaths(err: Error): Error {
    return this.maskFilePaths(err);
  }
  private removeSensitiveInformation() {
    delete this.telemetryClient.context.tags["ai.cloud.roleInstance"]; // cloud name
    delete this.telemetryClient.context.tags["ai.device.id"]; // machine name
    delete this.telemetryClient.context.tags["ai.user.accountId"]; // subscription
  }

  private telemetryOptedIn(): boolean {
    return this.telemetryObject.telemetryEnabled ;
  }
  private maskFilePaths(err: Error): Error {
    try {
      const regex = /\/(.*)\//gmi;
      const regex2 = /\w:\\(?:[^\\\s]+\\)+/gmi;
      err.message = err.message.replace(regex, "");
      err.stack = err.stack.replace(regex, "");
      err.stack = err.stack.replace(regex2, "");
      return err;
    } catch (err) {
      this.reportError("maskFilePaths", err);
    }
  }

}
