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
  private m_telemetryOptIn = true;//change to false once done
  private m_telemetryClient = appInsights.defaultClient;
  private m_appInsights: any;
  private m_telemetrySource = "";
  private m_testData: any;
  private events_sent = 0;
  private exceptions_sent = 0;


  constructor(instrumentationKey: string, telemetryType2: telemetryType, testData = false) {
    //checks to make sure it only displays the opt-in message once
    this.m_instrumentationKey = instrumentationKey;
    this.m_testData = testData;
    //declaring telemetry structure
    if (telemetryType2.toLowerCase() === 'applicationinsights') {
      this.m_telemetrySource = telemetryType.applicationinsights;
    } else if (telemetryType2.toLowerCase() === 'oteljs') {
      this.m_telemetrySource = telemetryType.OtelJs;
    }


    if (!this.m_testData && this.checkPrompt()) {
      this.telemetryOptIn();
    }
    appInsights.setup(this.m_instrumentationKey)
      .setAutoCollectConsole(true)
      .setAutoCollectExceptions(false)
      .setAutoCollectDependencies(true)
      .start()
    this.m_telemetryClient = appInsights.defaultClient;
    this.removeSensitiveInformation();
  }
  public async reportEvent(eventName: string, data: object, elapsedTime = 0): Promise<void> {
    if (this.telemetryOptedIn()) {
      if (this.m_telemetrySource === telemetryType.applicationinsights) {
        this.reportEventApplicationInsights(eventName, data);
      } else if (this.m_telemetrySource === telemetryType.OtelJs) {
      }
    }
  }
  public async reportEventApplicationInsights(eventName: string, data: object, elapsedTime = 0): Promise<void> {
    if (this.telemetryOptedIn()) {
      for (let [key, value] of Object.entries(data)) {
        try {
          if (!this.m_testData) {
            this.m_telemetryClient.trackEvent({ name: eventName, properties: { [key]: value } });
          }
          this.events_sent++;
        } catch (err) {
          this.reportError("sendTelemetryEvents", err);
        }
      }
    }
  }
  public async reportError(eventName: string, err: Error): Promise<void> {
    if (this.telemetryOptedIn()) {
      if (this.telemetryOptedIn()) {
        if (this.m_telemetrySource === telemetryType.applicationinsights) {
          this.reportErrorApplicationInsights(eventName, err);
        } else if (this.m_telemetrySource === telemetryType.OtelJs) {

        }
      }
    }
  }
  public async reportErrorApplicationInsights(eventName: string, err: Error): Promise<void> {
    if (this.telemetryOptedIn()) {
      err.name = eventName;
      this.m_telemetryClient.trackException({ exception: this.maskFilePaths(err) });
      this.exceptions_sent++;
    }
  }
  public addTelemetry(data: { [k: string]: any }, key: string, value: any): object {
    data[key] = value;
    return data;
  }


  public checkPrompt(fileLocation = "/officeAddinTelemetry.txt"): boolean {

    const path = require('os').homedir() + fileLocation;
    // checks to see if file exists
    try {
      if (fs.existsSync(path)) {
        var text = fs.readFileSync(path, "utf8");
        if (text.includes(this.getTelemtryKey())) {
          return false;
        } else {
          fs.writeFileSync(path, this.getTelemtryKey());
          return true;
        }
      } else {
        fs.writeFile(path, this.getTelemtryKey(), (err) => {
          if (err) console.log(err);

          return true;
        });
      }
    } catch (err) {
      this.reportError("checkPrompt", err);
    }
    return true;
  }

  public telemetryOptIn(testValue = 0): void {
    var chalk = require('chalk');
    var readlineSync = require('readline-sync');
    var response = "";
    var question = chalk.blue('-----------------------------------------\nDo you want to opt in for telemetry?[y/n]\n-----------------------------------------')
    if (testValue === 0) {
      try {
        response = readlineSync.question(question);
      } catch (err) {
        this.reportError("TelemetryOptIn", err)
      }
    }
    if (response.toLowerCase() === 'y' || testValue === 1) {
      this.m_telemetryOptIn = true;
      console.log('Telemetry will be sent!');
    } else {
      this.m_telemetryOptIn = false;
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
    return this.events_sent;
  }
  public getExceptionsSent(): any {
    return this.exceptions_sent;
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
    } catch(err) {
      this.reportError("maskFilePaths", err);
    }
  }

}
