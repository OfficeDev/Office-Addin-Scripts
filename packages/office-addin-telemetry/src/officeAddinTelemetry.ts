import * as appInsights from "applicationinsights";
import * as chalk from "chalk";
import * as os from "os";
import * as path from "path";
import * as readLine from "readline-sync";
import * as jsonData from "./telemetryJsonData";
/**
 * Specifies the telemetry infrastructure the user wishes to use
 * @enum Application Insights: Microsoft Azure service used to collect and query through data
 */
export enum telemetryType {
  applicationinsights = "applicationInsights",
}
/**
 * Level controlling what type of telemetry is being sent
 * @enum basic: basic level of telemetry, only sends errors
 * @enum verbose: verbose level of telemetry, sends errors and events
 */
export enum telemetryLevel {
  basic = "basic",
  verbose = "verbose",
}

const telemetryJsonFilePath: string = path.join(os.homedir(), "/officeAddinTelemetry.json");

/**
 * Telemetry object necessary for initialization of telemetry package
 * @member groupName Telemetry Group name that will be written to the telemetry config file (i.e. telemetryJsonFilePath)
 * @member projectName The name of the project that is using the telemetry package (e.g "generator-office")
 * @member instrumentationKey Instrumentation key for telemetry resource
 * @member promptQuestion Question displayed to user over opt-in for telemetry
 * @member raisePrompt Specifies whether to raise telemetry prompt (this allows for using a custom prompt)
 * @member telemetryLevel User's response to the prompt for telemetry and level of telemetry being sent
 * @member telemetryJsonFilePath Path to where telemetry json config file is written to.
 * @member telemetryType Telemetry infrastructure to send data
 * @member testData Allows user to run program without sending actual data
 */
export interface ITelemetryObject {
  groupName: string;
  projectName: string;
  instrumentationKey: string;
  promptQuestion: string;
  raisePrompt: boolean;
  telemetryJsonFilePath: string;
  telemetryLevel: telemetryLevel;
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
  private telemetryObject: ITelemetryObject;

  constructor(telemetryObj: ITelemetryObject) {
    try {
      this.telemetryObject = telemetryObj;

      if (this.telemetryObject.instrumentationKey === undefined) {
        throw new Error(chalk.default.red("Instrumentation not defined - cannot create telemetry object"));
      }

      if (this.telemetryObject.promptQuestion === undefined) {
        this.telemetryObject.promptQuestion = `Help improve ${this.telemetryObject.projectName} by allowing the collection of usage data. Would you like to particpate? Y/N`;
      }

      if (this.telemetryObject.telemetryJsonFilePath === undefined) {
        this.telemetryObject.telemetryJsonFilePath = telemetryJsonFilePath;
      }

      if (!this.telemetryObject.testData && this.telemetryObject.raisePrompt && jsonData.promptForTelemetry(this.telemetryObject.groupName, this.telemetryObject.telemetryJsonFilePath)) {
        this.telemetryOptIn();
      } else {
        jsonData.writeTelemetryJsonData(this.telemetryObject.groupName, this.telemetryObject.telemetryLevel, this.telemetryObject.telemetryJsonFilePath);
      }

      appInsights.setup(this.telemetryObject.instrumentationKey)
        .setAutoCollectExceptions(false)
        .start();
      this.telemetryClient = appInsights.defaultClient;
      this.removeApplicationInsightsSensitiveInformation();
    } catch (err) {
      console.log(`Failed to create telemetry object.\n${err}`);
    }
  }

  /**
   * Reports custom event object to telemetry structure
   * @param eventName Event name sent to telemetry structure
   * @param data Data object sent to telemetry structure
   * @param timeElapsed Optional parameter for custom metric in data object
   */
  public async reportEvent(eventName: string, data: object, timeElapsed = 0): Promise<void> {
    if (this.telemetryLevel() === telemetryLevel.verbose) {
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
    if (this.telemetryLevel() === telemetryLevel.verbose) {
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
    this.telemetryClient.trackException({ exception: this.maskFilePaths(err) });
    this.exceptionsSent++;
  }

  /**
   * Adds key and value(s) to given object
   * @param data Object used to contain custom event data
   * @param key Name of custom event data collected
   * @param value Data the user wishes to send
   * @param elapsedTime Optional duration of time for data to be collected
   * @returns The updated object with the new telemetry event added
   */
  public addTelemetry(data: { [k: string]: any }, key: string, value: any, elapsedTime: any = 0): object {
    data[key] = { value, elapsedTime };
    return data;
  }

  /**
   * Deletes specified key and value(s) from given object
   * @param data Object used to contain custom event data
   * @param key Name of key that is deleted along with corresponding values
   * @returns The updated object with the telemetry event removed
   */
  public deleteTelemetry(data: { [k: string]: any }, key: string): object {
    delete data[key];
    return data;
  }

  /**
   * Prompts user for telemetry participation once and records response
   * @param testData Specifies whether test code is calling this method
   * @param testReponse Specifies test response
   */
  public telemetryOptIn(testData: boolean = this.telemetryObject.testData, testResponse: string = ""): void {
    try {
      let response: string = "";
      if (testData) {
        response = testResponse;
      } else {
        response = readLine.question(chalk.default.blue(`${this.telemetryObject.promptQuestion}\n`));
      }
      if (response.toLowerCase() === "y") {
        this.telemetryObject.telemetryLevel = telemetryLevel.verbose;
      } else {
        this.telemetryObject.telemetryLevel = telemetryLevel.basic;
      }
      jsonData.writeTelemetryJsonData(this.telemetryObject.groupName, this.telemetryObject.telemetryLevel, this.telemetryObject.telemetryJsonFilePath);
      if (!this.telemetryObject.testData) {
        console.log(chalk.default.green(this.telemetryObject.telemetryLevel ? "Telemetry will be sent!" : "You will not be sending telemetry"));
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
   * @returns Whether telemetry is turned on or off
   */
  public isTelemetryOn(): boolean {
    return appInsights.defaultClient.config.samplingPercentage === 100;
  }

  /**
   * Returns the instrumentation key associated with the resource
   * @returns The telemetry instrumentation key
   */
  public getTelemetryKey(): string {
    return this.telemetryObject.instrumentationKey;
  }

  /**
   * Returns the amount of events that have been sent
   * @returns The count of events sent
   */
  public getEventsSent(): any {
    return this.eventsSent;
  }

  /**
   * Returns the amount of exceptions that have been sent
   * @returns The count of exceptions sent
   */
  public getExceptionsSent(): any {
    return this.exceptionsSent;
  }

  /**
   * Returns the telemetry level of the current object
   * @returns Telemetry level of the object
   */
  public telemetryLevel(): string {
    return this.telemetryObject.telemetryLevel;
  }

  /**
   * Returns parsed file path, scrubbing file names and sensitive information
   * @returns Error after removing PII
   */
  public maskFilePaths(err: Error): Error {
    try {
      const regexRemoveUserFilePaths = /\/(.*)\//gmi;
      const regexRemoveUserFilePathsFromStack = /\w:\\(?:[^\\\s]+\\)+/gmi;
      err.message = err.message.replace(regexRemoveUserFilePaths, "");
      err.stack = err.stack.replace(regexRemoveUserFilePaths, "");
      err.stack = err.stack.replace(regexRemoveUserFilePathsFromStack, "");
      return err;
    } catch (err) {
      this.reportError("maskFilePaths", err);
    }
  }
  /**
   * Removes sensitive information fields from ApplicationInsights data
   */
  private removeApplicationInsightsSensitiveInformation() {
    delete this.telemetryClient.context.tags["ai.cloud.roleInstance"]; // cloud name
    delete this.telemetryClient.context.tags["ai.device.id"]; // machine name
    delete this.telemetryClient.context.tags["ai.user.accountId"]; // subscription
  }
}
