import * as appInsights from "applicationinsights";

import { Base, ExceptionData, ExceptionDetails, StackFrame } from "applicationinsights/out/Declarations/Contracts";
import * as fs from "fs";
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

export class OfficeAddinTelemetry {
  public  chalk = require("chalk");
  private telemetryClient = appInsights.defaultClient;
  private telemetrySource = "";
  private eventsSent = 0;
  private exceptionsSent = 0;
  private path = require("os").homedir() + "/officeAddinTelemetry.txt";
  private telemetryObject;

  constructor(telemetryObject1: telemetryObject, groupName = undefined) {
    // checks to make sure it only displays the opt-in message once
    if (groupName === undefined){
    this.telemetryObject =  {
      groupName: telemetryObject1.groupName,
      instrumentationKey: telemetryObject1.instrumentationKey,
      promptQuestion: telemetryObject1.promptQuestion,
      telemetryEnabled: telemetryObject1.telemetryEnabled,
      telemetryType: telemetryObject1.telemetryType,
      testData: telemetryObject1.testData,
    }
  } else {
    this.copyInfo(groupName);
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
    this.path = require("os").homedir() + "/officeAddinTelemetry.txt";
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
  public deleteTelemetry(data: { [k: string]: any}, key: string, value: any, elapsedTime: any = 0): object {
    delete data[key];
    return data;
  }

  public checkPrompt(newFileLocation = this.path): boolean {
    try {
      this.path = newFileLocation;
      const name = this.telemetryObject.groupName;
      const obj = {};
      obj[name] =  this.telemetryObject.groupName;
      if (fs.existsSync(this.path)) {
        const text = fs.readFileSync(this.path, "utf8");
        if (text.includes(this.telemetryObject.groupName)) {
          const old = text.substring(text.indexOf(this.telemetryObject.groupName, 0), text.indexOf("}", text.indexOf(this.telemetryObject.groupName))+ 1);
          if(old.includes('"telemetryEnabled": false,')){
            this.telemetryObject.telemetryEnabled  = false;
          }
          return false;
        } else {
          fs.writeFileSync(this.path, text + JSON.stringify({ [name]: this.telemetryObject }, null, 2));
          return true;
        }
      } else {
        fs.writeFileSync(this.path, JSON.stringify({ [name]: this.telemetryObject }, null, 2));
          return true;
      }
    } catch (err) {
      this.reportError("checkPrompt", err);
    }
    return true;
  }

  public copyInfo(groupName){
        try {
      const name = groupName;
      const obj = {};
      obj[name] =  groupName;
      if (fs.existsSync(this.path)) {
        const text = fs.readFileSync(this.path, "utf8");
        if (text.includes(groupName)) {
          const data = text.substring(text.indexOf(groupName, 0) + 3 + groupName.length, text.indexOf("}", text.indexOf(groupName))+ 1);
          this.telemetryObject = JSON.parse(data);
        }
      } else {
        console.log("Group does not exist!");
        throw new Error();
      }
    } catch (err) {
      console.log(this.chalk.red("Group does not exist!"));
      this.reportError("copyInfo", err);
    }
  }
  public telemetryOptIn(testValue = 0): void {
    const readlineSync = require("readline-sync");
    const chalk = require("chalk");
    if (this.telemetryObject.promptQuestion === "") {
      this.telemetryObject.telemetryEnabled = false;
    } else {
    if (testValue === 0) {
    var text = fs.readFileSync(this.path, "utf8");
    }
    let response = "";
    if (testValue === 0) {
      try {
        response = readlineSync.question( chalk.blue(this.telemetryObject.promptQuestion));
      } catch (err) {
        this.reportError("TelemetryOptIn", err);
      }
    }
    if (testValue === 1 || response.toLowerCase() === "y") {
      this.telemetryObject.telemetryEnabled = true;
      console.log(chalk.green("Telemetry will be sent!"));
    } else {
      this.telemetryObject.telemetryEnabled = false;
      if (testValue !== 2 ) {
      const old = text.substring(text.indexOf(this.telemetryObject.groupName, 0), text.indexOf("}", text.indexOf(this.telemetryObject.groupName))+ 1);
      const updated = old.replace('"telemetryEnabled": true,', '"telemetryEnabled": false,');
      fs.writeFileSync(this.path, text.replace(old,updated));
      }
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
  public getTelemtryKey(): string {
    return this.telemetryObject.instrumentationKey;
  }
  public getTelemtryObject(): Object {
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
