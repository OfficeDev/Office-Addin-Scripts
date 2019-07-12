import * as appInsights from "applicationinsights";

import { Base, ExceptionData, ExceptionDetails, StackFrame } from 'applicationinsights/out/Declarations/Contracts';
import { inherits } from 'util';
import * as readline from "readline";//used
import * as fs from 'fs';//used
import * as chalk from 'chalk';//used
import { getMaxListeners } from "cluster";
export enum telemetryType {
  applicationinsights = "applicationInsights",
  OtelJs = "OtelJs",
}
export class OfficeAddinTelemetry {
  private m_instrumentationKey: string = "";
  private m_telemetryOptIn = true;
  private m_telemetryClient = appInsights.defaultClient;
  private m_telemetrySource = "";
  private m_testData: any;
  private m_linkpackages = "";
  private m_events_sent = 0;
  private m_exceptions_sent = 0;
  private m_path = require('os').homedir() + "/officeAddinTelemetry.txt";
  constructor(instrumentationKey: string, telemetryType2: telemetryType, linkpackages ="", testData = false) {
    //checks to make sure it only displays the opt-in message once
    this.m_instrumentationKey = instrumentationKey;
    this.m_testData = testData;
    this.m_linkpackages = linkpackages;
    //declaring telemetry structure
    if (telemetryType2 === telemetryType.applicationinsights) {
      this.m_telemetrySource = telemetryType.applicationinsights;
    } else if (telemetryType2.toLowerCase() === 'oteljs') {
      this.m_telemetrySource = telemetryType.OtelJs;
    }

    if (!this.m_testData && this.checkPrompt()) {
      this.telemetryOptIn();
    }
    this.m_path = require('os').homedir() + "/officeAddinTelemetry.txt";
    appInsights.setup(this.m_instrumentationKey)
      .setAutoCollectConsole(true)
      .setAutoCollectExceptions(false)
      .start()
    this.m_telemetryClient = appInsights.defaultClient;
    this.removeSensitiveInformation();
  }
  public async reportEvent(eventName: string, data: object,timeElapsed = 0): Promise<void> {
    if (this.telemetryOptedIn() && !this.m_testData) {
      if (this.m_telemetrySource === telemetryType.applicationinsights) {
        this.reportEventApplicationInsights(eventName, data);
      } else if (this.m_telemetrySource === telemetryType.OtelJs) {
        //to be implemented in a later date
      }
    }
  }
  public async reportEventApplicationInsights(eventName: string, data: object): Promise<void> {
    if (this.telemetryOptedIn()) {
      for (let [key, {value,elapsedTime}] of Object.entries(data)) {
        try {
          if (!this.m_testData) {
            var temp = "DurationFor" + [key].toLocaleString();
            this.m_telemetryClient.trackEvent({ name: eventName, properties: { [key]: value}, measurements: { temp: elapsedTime}});
            //console.log({ name: eventName, properties: { [key]: value}, metrics: { [key]: elapsedTime} });
          }
          this.m_events_sent++;
        } catch (err) {
          this.reportError("sendTelemetryEvents", err);
        }
      }
    }
  }
  public async reportError(errorName: string, err: Error): Promise<void> {
    if (this.telemetryOptedIn()) {
        if (this.m_telemetrySource === telemetryType.applicationinsights) {
          this.reportErrorApplicationInsights(errorName, err);
        } else if (this.m_telemetrySource === telemetryType.OtelJs) {
        //to be implemented in a later date            
        }
    }
  }
  public async reportErrorApplicationInsights(eventName: string, err: Error): Promise<void> {
    if (this.telemetryOptedIn()) {
      err.name = eventName;
      if(this.m_testData){
        err.name = eventName;
      }
      this.m_telemetryClient.trackException({ exception: this.maskFilePaths(err) });
      this.m_exceptions_sent++;
    }
  }
  public addTelemetry(data: { [k: string]: any}, key: string, value: any, elapsedTime: any = 0): object {
    data[key] = {value,elapsedTime};
    return data;
  }


  public checkPrompt(newFileLocation = this.m_path): boolean {
    try {
      this.m_path = newFileLocation;
      if (fs.existsSync(this.m_path)) {
        var text = fs.readFileSync(this.m_path, "utf8");
        if (text.includes(this.getTelemtryKey()) || text.includes(this.m_linkpackages) && this.m_linkpackages !== "") {
          if(text.includes(this.getTelemtryKey() + "no") || text.includes(this.m_linkpackages + "no")){
            this.m_telemetryOptIn = false;
          }
          return false;
        } else {
          fs.writeFileSync(this.m_path, text + this.getTelemtryKey() + this.m_linkpackages);
          return true;
        }
      } else {
        fs.writeFileSync(this.m_path, this.getTelemtryKey() + this.m_linkpackages);
          return true;
      }
    } catch (err) {
      this.reportError("checkPrompt", err);
    }
    return true;
  }

  public telemetryOptIn(testValue = 0): void {
    var chalk = require('chalk');
    var readlineSync = require('readline-sync');
    if(testValue === 0){
    var text = fs.readFileSync(this.m_path, "utf8");
    }
    var response = "";
    var question = chalk.blue('-----------------------------------------\nDo you want to opt in for telemetry?[y/n]\n-----------------------------------------');
    if (testValue === 0) {
      try {
        response = readlineSync.question(question);
      } catch (err) {
        this.reportError("TelemetryOptIn", err);
      }
    }
    if (response.toLowerCase() === 'y' || testValue === 1) {
      this.m_telemetryOptIn = true;
      console.log('Telemetry will be sent!');
    } else {
      this.m_telemetryOptIn = false;
      fs.writeFileSync(this.m_path, text + "no");
      console.log('You will not be sending telemetry');
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
    return this.m_instrumentationKey;
  }
  public getEventsSent(): any {
    return this.m_events_sent;
  }
  public getExceptionsSent(): any {
    return this.m_exceptions_sent;
  }

  public telemetryOptedIn2(): boolean {//for mocha test
    return this.m_telemetryOptIn;
  }

  private removeSensitiveInformation() {
    delete this.m_telemetryClient.context.tags['ai.cloud.roleInstance'];//cloud name
    delete this.m_telemetryClient.context.tags['ai.device.id'];//machine name
    delete this.m_telemetryClient.context.tags['ai.user.accountId'];//subscription #
  }

  private telemetryOptedIn(): boolean {
    return this.m_telemetryOptIn;
  }
  public testMaskFilePaths(err: Error): Error {
    return this.maskFilePaths(err);
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
