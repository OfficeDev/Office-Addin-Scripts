import * as appInsights from "applicationinsights";
import * as chalk from "chalk";
import * as fs from "fs";
import * as os from "os";
import * as path from "path";
import * as readLine from "readline-sync";
export enum telemetryType {
  applicationinsights = "applicationInsights",
  // OtelJs = "OtelJs" - Not yet implemented
}

const telemetryJsonFilePath: string = path.join(os.homedir(), "/officeAddinTelemetry.json");

/**
 * Telemetry object necesary for initialization of telemetry package
 * @member groupName Event name sent to telemetry structure
 * @member instrumentationKey Instrumentation key for telemetry resource
 * @member promptQuestion Question displayed to User over opt-in for telemetry
 * @member telemetryEnabled User's response to the prompt for telemetry
 * @member telemetryType Telemetry infrastructure to send data
 * @member testData Allows user to run program without sending actuall data
 */
export interface telemetryObject {
  groupName: string;
  instrumentationKey: string;
  promptQuestion: string;
  raisePrompt: boolean;
  telemetryEnabled: boolean;
  telemetryJsonFilePath: string;
  telemetryType: telemetryType;
  testData: boolean;
}

/**
 * Creates and initializes member variables while prompting user for telemetry collection when necessary
 * @param telemetryObject
 */
export class OfficeAddinTelemetry {
  private telemetryClient = appInsights.defaultClient;
  private eventsSent: number = 0;
  private exceptionsSent: number = 0;
  private telemetryObject: telemetryObject;

  constructor(telemetryObj: telemetryObject) {
    try {
      this.telemetryObject = telemetryObj;
      if (this.telemetryObject.telemetryJsonFilePath === undefined){
        this.telemetryObject.telemetryJsonFilePath = telemetryJsonFilePath;
      }

      if (!this.telemetryObject.testData && this.telemetryObject.raisePrompt && promptForTelemetry(this.telemetryObject.groupName)) {
        this.telemetryOptIn();
      } else {
        const telemetryJsonData = readTelemetryJsonData(this.telemetryObject.telemetryJsonFilePath);
        if (telemetryJsonData) {
          if (!groupNameExists(telemetryJsonData, this.telemetryObject.groupName)) {
            telemetryJsonData.telemetryInstances[this.telemetryObject.groupName] = this.telemetryObject.telemetryEnabled;
            writeTelemetryJsonData(telemetryJsonData, this.telemetryObject.telemetryJsonFilePath);
          }
        } else {
          writeNewTelemetryJsonFile(this.telemetryObject.groupName, this.telemetryObject.telemetryEnabled, this.telemetryObject.telemetryJsonFilePath);
        }
      }

      if (this.telemetryObject.instrumentationKey === undefined) {
        throw new Error("Instrumentation not defined - cannot create telemetry object");
      }

      appInsights.setup(this.telemetryObject.instrumentationKey)
        .setAutoCollectConsole(true)
        .setAutoCollectExceptions(false)
        .start();
      this.telemetryClient = appInsights.defaultClient;
      this.removeSensitiveInformation();
    } catch (err) {
      console.log(`Failed to create telemetry object.\n${err}`);
    }
  }

  /**
   * Reports custom event object to telemetry structure
   * @param eventName Event name sent to telemetry structure
   * @param data Data object sent to telemetry structure
   * @param timeElapsed Optional parameter for custom metric in data object sent
   */
  public async reportEvent(eventName: string, data: object, timeElapsed = 0): Promise<void> {
    if (this.telemetryOptedIn()) {
      this.reportEventApplicationInsights(eventName, data);
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
      for (const [key, { value, elapsedTime }] of Object.entries(data)) {
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
    this.reportErrorApplicationInsights(errorName, err);
  }

  /**
   * Reports error to Application Insights
   * @param errorName Error name sent to Application Insights
   * @param err Error sent to Application Insights
   */
  public async reportErrorApplicationInsights(errorName: string, err: Error): Promise<void> {
    err.name = errorName;
    if (this.telemetryObject.testData) {
      err.name = errorName;
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
  * Prompts user for telemtry participation once and records response
  * @param mochaTest Specifies whether test code is calling this method
  * @param testReponse Specifies test response
  */
  public telemetryOptIn(mochaTest: boolean = false, testResponse: string = ""): void {
    try {
      let response: string = "";
      if (mochaTest) {
        response = testResponse;
      } else {
        response = readLine.question(chalk.default.blue(this.telemetryObject.promptQuestion));
      }

      const enableTelemetry = response.toLowerCase() === "y";
      const telemetryJsonData: any = readTelemetryJsonData(this.telemetryObject.telemetryJsonFilePath);

      if (telemetryJsonData) {
        this.telemetryObject.telemetryEnabled = enableTelemetry;
        telemetryJsonData.telemetryInstances[this.telemetryObject.groupName] = enableTelemetry;
        writeTelemetryJsonData(telemetryJsonData, this.telemetryObject.telemetryJsonFilePath);
        if (!this.telemetryObject.testData) {
          console.log(chalk.default.green(enableTelemetry ? "Telemetry will be sent!" : "You will not be sending telemetry"));
        }
      } else {
        writeNewTelemetryJsonFile(this.telemetryObject.groupName, enableTelemetry, this.telemetryObject.telemetryJsonFilePath);
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
   * Returns whether the telemetry is currently on or off
   */
  public isTelemetryOn(): boolean {
    if (appInsights.defaultClient.config.samplingPercentage === 100) {
      return true;
    } else {
      return false;
    }
  }

  /**
   * Returns the instrumentation key associated with the resource
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
  /**
   * Returns whether the user opted in or not
   */
  public telemetryOptedIn(): boolean {
    return this.telemetryObject.telemetryEnabled;
  }
  /**
   * Returns whether the user opted in or not
   */
  public maskFilePaths(err: Error): Error {
    try {
      const regexRemoveUserFilePaths   = /\/(.*)\//gmi;
      const regexRemoveUserFilePathsFromStack   = /\w:\\(?:[^\\\s]+\\)+/gmi;
      err.message = err.message.replace(regexRemoveUserFilePaths  , "");
      err.stack = err.stack.replace(regexRemoveUserFilePaths  , "");
      err.stack = err.stack.replace(regexRemoveUserFilePathsFromStack  , "");
      return err;
    } catch (err) {
      this.reportError("maskFilePaths", err);
    }
  }

  private removeSensitiveInformation() {
    delete this.telemetryClient.context.tags["ai.cloud.roleInstance"]; // cloud name
    delete this.telemetryClient.context.tags["ai.device.id"]; // machine name
    delete this.telemetryClient.context.tags["ai.user.accountId"]; // subscription
  }
}

/**
 * Allows developer to create prompts and responses in other applications before object creation
 * @param groupName Event name sent to telemetry structure
 * @param telemetryEnabled Whether user agreed to data collection
 */
export function promptForTelemetry(groupName: string, jsonFilePath: string = telemetryJsonFilePath): boolean {
  try {
    const jsonData: any = readTelemetryJsonData(jsonFilePath);
    if (jsonData) {
      return !groupNameExists(jsonData, groupName);
    }
    return true;
  } catch (err) {
    console.log(chalk.default.red(err));
  }
}

export function readTelemetryJsonData(jsonFilePath): any {
  if (fs.existsSync(jsonFilePath)) {
    const jsonData = fs.readFileSync(jsonFilePath, "utf8");
    return JSON.parse(jsonData.toString());
  }
}

export function writeTelemetryJsonData(jsonData: any, jsonFilePath) {
  fs.writeFileSync(jsonFilePath, JSON.stringify((jsonData), null, 2));
}

export function writeNewTelemetryJsonFile(groupName: string, telemetryEnabled, jsonFilePath: string) {
  let jsonData = {};
  jsonData[groupName] = telemetryEnabled;
  jsonData = { telemetryInstances: jsonData };
  writeTelemetryJsonData(jsonData, jsonFilePath);
}

export function groupNameExists(jsonData: any, groupName: string): boolean {
  return Object.getOwnPropertyNames(jsonData.telemetryInstances).includes(groupName);
}