import * as appInsights from "applicationinsights";
import * as chalk from "chalk";
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

/**
 * Telemetry object necessary for initialization of telemetry package
 * @member groupName Telemetry Group name that will be written to the telemetry config file (i.e. telemetryJsonFilePath)
 * @member projectName The name of the project that is using the telemetry package (e.g "generator-office")
 * @member instrumentationKey Instrumentation key for telemetry resource
 * @member promptQuestion Question displayed to user over opt-in for telemetry
 * @member raisePrompt Specifies whether to raise telemetry prompt (this allows for using a custom prompt)
 * @member telemetryLevel User's response to the prompt for telemetry and level of telemetry being sent
 * @member telemetryType Telemetry infrastructure to send data
 * @member testData Allows user to run program without sending actual data
 */
export interface ITelemetryOptions {
  groupName: string;
  projectName: string;
  instrumentationKey: string;
  promptQuestion: string;
  raisePrompt: boolean;
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
  private telemetryObject: ITelemetryOptions;

  constructor(telemetryObj: ITelemetryOptions) {
    try {
      this.telemetryObject = telemetryObj;

      if (this.telemetryObject.instrumentationKey === undefined) {
        throw new Error(chalk.default.red("Instrumentation not defined - cannot create telemetry object"));
      }

      if (this.telemetryObject.promptQuestion === undefined) {
        this.telemetryObject.promptQuestion = `Help improve ${this.telemetryObject.projectName} by allowing the collection of usage data. Would you like to particpate? Y/N`;
      }
      if (this.telemetryObject.groupName === undefined) {
        throw new Error(chalk.default.red("Group Name not defined - cannot create telemetry object"));
      }

      if (jsonData.groupNameExists(this.telemetryObject.groupName)) {
        this.telemetryObject.telemetryLevel = jsonData.readTelemetryLevel(this.telemetryObject.groupName);
      }

      if (!this.telemetryObject.testData && this.telemetryObject.raisePrompt && jsonData.needToPromptForTelemetry(this.telemetryObject.groupName)) {
        this.telemetryOptIn();
      } else {
        jsonData.writeTelemetryJsonData(this.telemetryObject.groupName, this.telemetryObject.telemetryLevel);
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
   * @param byPass Boolean to still report events even when telemetry level is set to basic
   */
  public async reportEvent(eventName: string, data: object, byPass: boolean = false): Promise<void> {
    if (this.telemetryLevel() === telemetryLevel.verbose || byPass === true) {
      this.reportEventApplicationInsights(eventName, data, byPass);
    }
  }

  /**
   * Reports custom event object to Application Insights
   * @param eventName Event name sent to Application Insights
   * @param data Data object sent to Application Insights
   * @param byPass Boolean to still report events to Application Insights even when telemetry level is set to basic
   */
  public async reportEventApplicationInsights(eventName: string, data: object, byPass: boolean = false): Promise<void> {
    if (this.telemetryLevel() === telemetryLevel.verbose || byPass === true) {
      const telemetryEvent = new appInsights.Contracts.EventData();
      telemetryEvent.name = eventName;
      try {
        for (const [key, [value, elapsedTime]] of Object.entries(data)) {
          telemetryEvent.properties[key] = value;
          telemetryEvent.measurements[key + " durationElapsed"] = elapsedTime;
        }
        if (!this.telemetryObject.testData) {
          this.telemetryClient.trackEvent(telemetryEvent);
        }
        this.eventsSent++;
      } catch (err) {
        this.reportError("sendTelemetryEvents", err);
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
      jsonData.writeTelemetryJsonData(this.telemetryObject.groupName, this.telemetryObject.telemetryLevel);
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
