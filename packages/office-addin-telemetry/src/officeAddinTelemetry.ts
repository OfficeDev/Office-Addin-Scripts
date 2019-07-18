import * as appInsights from "applicationinsights";
import * as chalk from "chalk";
import * as fs from "fs";
import * as os from "os";
import * as path from "path";
import * as readLine from "readline-sync";
export enum telemetryType {
  applicationinsights = "applicationInsights",
  OtelJs = "OtelJs",
}

const telemetryJsonFilePath: string = path.join(os.homedir(), "/officeAddinTelemetry.json");

/**
 * Telemetry object necesary for initialization of telemetry package
 * @param groupName Event name sent to telemetry structure
 * @param instrumentationKey Instrumentation key for telemetry resource
 * @param promptQuestion Question displayed to User over opt-in for telemetry
 * @param telemetryEnabled User's response to the prompt for telemetry
 * @param telemetryType Telemetry infrastructure to send data
 * @param testData Allows user to run program without sending actuall data
 */
export interface telemetryObject {
  groupName: string;
  instrumentationKey: string;
  promptQuestion: string;
  telemetryEnabled: boolean;
  telemetryType: telemetryType;
  testData: boolean;
}

/**
 * Allows developer to create a telemetry object and send custom events and exceptions to a unique telemetry structure
 */
export class OfficeAddinTelemetry {
  public chalk = require("chalk");
  private telemetryClient = appInsights.defaultClient;
  private telemetrySource = "";
  private eventsSent = 0;
  private exceptionsSent = 0;
  private path = require("os").homedir() + "/officeAddinTelemetry.json";
  private telemetryObject;
  /**
   * Creates and intializes memeber variables while prompting user for telemetry collection when necessary
   * @param telemetryObject
   */
  constructor(telemetryObj: telemetryObject) {
    // checks to make sure it only displays the opt-in message once
    try {
      this.telemetryObject = {
        groupName: telemetryObj.groupName,
        instrumentationKey: telemetryObj.instrumentationKey,
        promptQuestion: telemetryObj.promptQuestion,
        telemetryEnabled: telemetryObj.telemetryEnabled,
        telemetryType: telemetryObj.telemetryType,
        testData: telemetryObj.testData,
      }
    } catch (err) {
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

  /**
   * Reports custom event object to telemetry structure
   * @param eventName Event name sent to telemetry structure
   * @param data Data object sent to telemetry structure
   * @param timeElapsed Optional parameter for custom metric in data object sent
   */
  public async reportEvent(eventName: string, data: object, timeElapsed = 0): Promise<void> {
    if (this.telemetryOptedIn()) {
      if (this.telemetrySource === telemetryType.applicationinsights) {
        this.reportEventApplicationInsights(eventName, data);
      } else if (this.telemetrySource === telemetryType.OtelJs) {
        console.log(this.chalk.red("Feature has not been yet created."));
      }
    }
  }

  /**
   * Reports custom event object to Application Insights
   * @param eventName Event name sent to Application Insights
   * @param data Data object sent to Application Insights
   * @param timeElapsed Optional parameter for custom metric in data object sent to Application Insights
   */
  public async reportEventApplicationInsights(eventName: string, data: object): Promise<void> {
    if (this.telemetryOptedIn()) {
      for (let [key, { value, elapsedTime }] of Object.entries(data)) {
        try {
          if (!this.telemetryObject.testData) {
            this.telemetryClient.trackEvent({ name: eventName, properties: { [key]: value }, measurements: { DurationElapsed: elapsedTime } });
          }
          this.eventsSent++;
        } catch (err) {
          this.reportError("sendTelemetryEvents", err);
        }
      }
    }
  }

  /**
   * Reports error to telemetry structure
   * @param errorName Error name sent to telemetry structure
   * @param err Error sent to telemetry structure
   */
  public async reportError(errorName: string, err: Error): Promise<void> {
    if (this.telemetrySource === telemetryType.applicationinsights) {
      this.reportErrorApplicationInsights(errorName, err);
    } else if (this.telemetrySource === telemetryType.OtelJs) {
      console.log(this.chalk.red("Feature has not been yet created."));
    }
  }

  /**
   * Reports error to Application Insights
   * @param errorName Error name sent to Application Insights
   * @param err Error sent to Application Insights
   */
  public async reportErrorApplicationInsights(eventName: string, err: Error): Promise<void> {
    err.name = eventName;
    if (this.telemetryObject.testData) {
      err.name = eventName;
    }
    this.telemetryClient.trackException({ exception: this.maskFilePaths(err) });
    this.exceptionsSent++;
  }

  /**
   * Adds key and value(s) to given object
   * @param data Object used to contain custom event data
   * @param key Name of custom event data collected
   * @param value Data the user wishes to send
   * @param elapsedTime Optional duration of time for data to be collected
   */
  public addTelemetry(data: { [k: string]: any }, key: string, value: any, elapsedTime: any = 0): object {
    data[key] = { value, elapsedTime };
    return data;
  }

  /**
   * Deletes specified key and value(s) from given object
   * @param data Object used to contain custom event data
   * @param key Name of key that is deleted along with corresponding values
   */
  public deleteTelemetry(data: { [k: string]: any }, key: string): object {
    delete data[key];
    return data;
  }

  /**
   * Checks whether user has been prompted before
   * @param newFileLocation Optional file location to specify the location of json file to write the user's group Name and response to prompt
   */
  public checkPrompt(newFileLocation = this.path): boolean {
    try {
      this.path = newFileLocation;
      const name = this.telemetryObject.groupName;
      if (fs.existsSync(this.path)) {
        const jsonData = fs.readFileSync(this.path, "utf8");
        if (jsonData !== "") {
          const projectJsonData = JSON.parse(jsonData.toString());
          if (Object.getOwnPropertyNames(projectJsonData.telemetryInstances).includes(this.telemetryObject.groupName)) {
            if (projectJsonData.telemetryInstances[name] === false) {
              this.telemetryObject.telemetryEnabled = false;
            } else {
              this.telemetryObject.telemetryEnabled = true;
            }
            return false;
          } else {
            projectJsonData.telemetryInstances[name] = this.telemetryObject.telemetryEnabled;
            fs.writeFileSync(this.path, JSON.stringify((projectJsonData), null, 2));
            return true;
          }
        } else {
          const projectJsonData = {};
          projectJsonData[name] = this.telemetryObject.telemetryEnabled;
          fs.writeFileSync(this.path, JSON.stringify(({ "telemetryInstances": projectJsonData }), null, 2));
        }
      } else {
        const projectJsonData = {};
        projectJsonData[name] = this.telemetryObject.telemetryEnabled;
        fs.writeFileSync(this.path, JSON.stringify(({ "telemetryInstances": projectJsonData }), null, 2));
      }
    } catch (err) {
      this.reportError("checkPrompt", err);
    }

    return true;
  }

  // /**
  //  * Prompts user for telemtry participation once and records response
  //  */
  // public telemetryOptIn(testValue = false): void {
  //   const readlineSync = require("readline-sync");
  //   const chalk = require("chalk");
  //   if (this.telemetryObject.promptQuestion === undefined || this.telemetryObject.promptQuestion === "") {
  //     this.telemetryObject.telemetryEnabled = false;
  //   } else {
  //     const name = this.telemetryObject.groupName;
  //     const jsonData = fs.readFileSync(this.path, "utf8");
  //     const projectJsonData = JSON.parse(jsonData.toString());
  //     let response = "";
  //     if (testValue === 0) {
  //       try {
  //         response = readlineSync.question(chalk.blue(this.telemetryObject.promptQuestion));
  //       } catch (err) {
  //         this.reportError("TelemetryOptIn", err);
  //       }
  //     }
  //     if (response.toLowerCase() === "y" || testValue === 1) {
  //       this.telemetryObject.telemetryEnabled = true;
  //       projectJsonData.telemetryInstances[name] = true;
  //       fs.writeFileSync(this.path, JSON.stringify((projectJsonData), null, 2));
  //       console.log(chalk.green("Telemetry will be sent!"));
  //     } else {
  //       this.telemetryObject.telemetryEnabled = false;
  //       projectJsonData.telemetryInstances[name] = false;
  //       fs.writeFileSync(this.path, JSON.stringify((projectJsonData), null, 2));
  //       console.log(chalk.red("You will not be sending telemetry"));
  //     }
  //   }
  // }

   /**
   * Prompts user for telemtry participation once and records response
   */
  public telemetryOptIn(mochaTest: boolean = false): void {
    try {
      if (promptForTelemetry(this.telemetryObject.groupName) && !mochaTest) {
        const response = readLine.question(chalk.default.blue(this.telemetryObject.promptQuestion));
        const telemetryJsonData: any = readTelemetryJsonData();
        const enableTelemetry = response.toLowerCase() === "y";

        if (telemetryJsonData) {
            this.telemetryObject.telemetryEnabled = enableTelemetry;
            telemetryJsonData.telemetryInstances[this.telemetryObject.groupName] = enableTelemetry;
            fs.writeFileSync(this.path, JSON.stringify((telemetryJsonData), null, 2));
            console.log(chalk.default.green(enableTelemetry ? "Telemetry will be sent!" : "You will not be sending telemetry"));
       } else {
            const projectJsonData = {};
            projectJsonData[this.telemetryObject.groupName] =  enableTelemetry;
            fs.writeFileSync(this.path, JSON.stringify(({"telemetryInstances": projectJsonData}), null, 2));
          }
      }
    } catch (err) {
      this.reportError("TelemetryOptIn", err);
    }
  }

  /**
   * Stops telemetry from being sent, by default telemetry will be on
   */
  public setTelemetryOff() {
    appInsights.defaultClient.config.samplingPercentage = 0;
  }

  /**
   * Starts sending telemetry, by default telemetry will be on
   */
  public setTelemetryOn() {
    appInsights.defaultClient.config.samplingPercentage = 100;
  }

  /**
   * Returns wheter telemetry is on(true) or off(false)
   */
  public isTelemetryOn(): boolean {
    if (appInsights.defaultClient.config.samplingPercentage === 100) {
      return true;
    } else {
      return false;
    }
  }

  /**
   * Returns wheter telemetry is on(true) or off(false)
   */
  public getTelemetryKey(): string {
    return this.telemetryObject.instrumentationKey;
  }

  /**
   * Returns amount of events that have been sent
   */
  public getEventsSent(): any {
    return this.eventsSent;
  }

  /**
   * Returns amount of exceptions that have been sent
   */
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
    return this.telemetryObject.telemetryEnabled;
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

/**
 * Allows developer to create prompts and responses in other applications before object creation
 * @param groupName Event name sent to telemetry structure
 * @param telemetryEnabled Whether user agreed to data collection
 */
export function promptForTelemetry(groupName: string) : boolean {
  try {
    const jsonData: any = readTelemetryJsonData();
    if (jsonData) {
      if (Object.getOwnPropertyNames(jsonData.telemetryInstances).includes(groupName)) {
        return false;
      }
      return true;
     }
     return true;
  } catch (err) {
    console.log(chalk.default.red(err));
  }
}

function readTelemetryJsonData() : any {
  if (fs.existsSync(telemetryJsonFilePath)) {
    const jsonData = fs.readFileSync(telemetryJsonFilePath, "utf8");
    return JSON.parse(jsonData.toString());
  }
}
