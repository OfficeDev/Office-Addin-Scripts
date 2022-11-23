// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as appInsights from "applicationinsights";
import * as readLine from "readline-sync";
import * as jsonData from "./usageDataSettings";
import * as defaults from "./defaults";

/* global process */

/**
 * Specifies the usage data infrastructure the user wishes to use
 * @enum Application Insights: Microsoft Azure service used to collect and query through data
 */
export enum UsageDataReportingMethod {
  applicationInsights = "applicationInsights",
}
/**
 * Level controlling what type of usage data is being sent
 * @enum off: off level of usage data, sends no usage data
 * @enum on: on level of usage data, sends errors and events
 */
export enum UsageDataLevel {
  off = "off",
  on = "on",
}

/**
 * Defines an error that is expected to happen given some situation
 * @member message Message to be logged in the error
 */
export class ExpectedError extends Error {
  constructor(message: string | undefined) {
    super(message);

    // need to adjust the prototype after super()
    // See https://github.com/Microsoft/TypeScript-wiki/blob/master/Breaking-Changes.md#extending-built-ins-like-error-array-and-map-may-no-longer-work
    Object.setPrototypeOf(this, ExpectedError.prototype);
  }
}

/**
 * UpdateData options
 * @member groupName Group name for usage data settings (Optional)
 * @member projectName The name of the project that is using the usage data package.
 * @member instrumentationKey Instrumentation key for usage data resource
 * @member promptQuestion Question displayed to user over opt-in for usage data (Optional)
 * @member raisePrompt Specifies whether to raise usage data prompt (this allows for using a custom prompt) (Optional)
 * @member usageDataLevel User's response to the prompt for usage data (Optional)
 * @member method The desired method to use for reporting usage data. (Optional)
 * @member isForTesting True if the data is just for testing, false for actual data that should be reported. (Optional)
 */
export interface IUsageDataOptions {
  groupName?: string;
  projectName: string;
  instrumentationKey: string;
  promptQuestion?: string;
  raisePrompt?: boolean;
  usageDataLevel?: UsageDataLevel;
  method?: UsageDataReportingMethod;
  isForTesting?: boolean;
  deviceID?: string;
}

/**
 * Creates and initializes member variables while prompting user for usage data collection when necessary
 * @param usageDataObject
 */
export class OfficeAddinUsageData {
  private usageDataClient = appInsights.defaultClient;
  private eventsSent: number = 0;
  private exceptionsSent: number = 0;
  private options: IUsageDataOptions;
  private defaultData = {
    Platform: process.platform,
    NodeVersion: process.version,
  };

  constructor(usageDataOptions: IUsageDataOptions) {
    try {
      this.options = {
        groupName: defaults.groupName,
        promptQuestion: "",
        raisePrompt: true,
        usageDataLevel: UsageDataLevel.off,
        method: UsageDataReportingMethod.applicationInsights,
        isForTesting: false,
        ...usageDataOptions,
      };

      if (this.options.instrumentationKey === undefined) {
        throw new Error("Instrumentation Key not defined - cannot create usage data object");
      }

      if (this.options.groupName === undefined) {
        throw new Error("Group Name not defined - cannot create usage data object");
      }

      if (jsonData.groupNameExists(this.options.groupName)) {
        this.options.usageDataLevel = jsonData.readUsageDataLevel(this.options.groupName);
      }

      // Generator-office will not raise a prompt because the yeoman generator creates the prompt.  If the projectName
      // is defaults.generatorOffice and a office-addin-usage-data file hasn't been written yet, write one out.
      if (
        this.options.projectName === defaults.generatorOffice &&
        this.options.instrumentationKey === defaults.instrumentationKeyForOfficeAddinCLITools &&
        jsonData.needToPromptForUsageData(this.options.groupName)
      ) {
        jsonData.writeUsageDataJsonData(this.options.groupName, this.options.usageDataLevel);
      }

      if (
        !this.options.isForTesting &&
        this.options.raisePrompt &&
        jsonData.needToPromptForUsageData(this.options.groupName)
      ) {
        this.usageDataOptIn();
      }
      this.options.deviceID = jsonData.readDeviceID();
      if (this.options.usageDataLevel === UsageDataLevel.on) {
        appInsights.setup(this.options.instrumentationKey).setAutoCollectExceptions(false).start();
        this.usageDataClient = appInsights.defaultClient;
        this.removeApplicationInsightsSensitiveInformation();
      }
    } catch (err) {
      throw new Error(err);
    }
  }

  /**
   * Reports custom event object to usage data structure
   * @param eventName Event name sent to usage data structure
   * @param data Data object sent to usage data structure
   */
  public async reportEvent(eventName: string, data: object): Promise<void> {
    if (this.getUsageDataLevel() === UsageDataLevel.on) {
      this.reportEventApplicationInsights(eventName, data);
    }
  }

  /**
   * Reports custom event object to Application Insights
   * @param eventName Event name sent to Application Insights
   * @param data Data object sent to Application Insights
   */
  public async reportEventApplicationInsights(eventName: string, data: object): Promise<void> {
    if (this.getUsageDataLevel() === UsageDataLevel.on) {
      const usageDataEvent = new appInsights.Contracts.EventData();
      usageDataEvent.name = this.options.isForTesting ? `${eventName}-test` : eventName;
      try {
        for (const [key, [value, elapsedTime]] of Object.entries(data)) {
          usageDataEvent.properties[key] = value;
          usageDataEvent.measurements[key + " durationElapsed"] = elapsedTime;
        }
        usageDataEvent.properties["deviceID"] = this.options.deviceID;
        this.usageDataClient.trackEvent(usageDataEvent);
        this.eventsSent++;
      } catch (err) {
        this.reportError("sendUsageDataEvents", err);
        throw new Error(err);
      }
    }
  }

  /**
   * Reports error to usage data structure
   * @param errorName Error name sent to usage data structure
   * @param err Error sent to usage data structure
   */
  public async reportError(errorName: string, err: Error): Promise<void> {
    if (this.getUsageDataLevel() === UsageDataLevel.on) {
      this.reportErrorApplicationInsights(errorName, err);
    }
  }

  /**
   * Reports error to Application Insights
   * @param errorName Error name sent to Application Insights
   * @param err Error sent to Application Insights
   */
  public async reportErrorApplicationInsights(errorName: string, err: Error): Promise<void> {
    if (this.getUsageDataLevel() === UsageDataLevel.on) {
      let error = Object.create(err);

      error.name = this.options.isForTesting ? `${errorName}-test` : errorName;
      this.usageDataClient.trackException({
        exception: this.maskFilePaths(error),
      });
      this.exceptionsSent++;
    }
  }

  /**
   * Prompts user for usage data participation once and records response
   * @param testData Specifies whether test code is calling this method
   * @param testReponse Specifies test response
   */
  public usageDataOptIn(testData: boolean = this.options.isForTesting, testResponse: string = ""): void {
    try {
      let response: string = "";
      if (testData) {
        response = testResponse;
      } else {
        response = readLine.question(`${this.options.promptQuestion}\n`);
      }
      if (response.toLowerCase() === "y") {
        this.options.usageDataLevel = UsageDataLevel.on;
      } else {
        this.options.usageDataLevel = UsageDataLevel.off;
      }
      jsonData.writeUsageDataJsonData(this.options.groupName, this.options.usageDataLevel);
    } catch (err) {
      this.reportError("UsageDataOptIn", err);
      throw new Error(err);
    }
  }

  /**
   * Stops usage data from being sent, by default usage data will be on
   */
  public setUsageDataOff() {
    appInsights.defaultClient.config.samplingPercentage = 0;
  }

  /**
   * Starts sending usage data, by default usage data will be on
   */
  public setUsageDataOn() {
    appInsights.defaultClient.config.samplingPercentage = 100;
  }

  /**
   * Returns whether the usage data is currently on or off
   * @returns Whether usage data is turned on or off
   */
  public isUsageDataOn(): boolean {
    return appInsights.defaultClient.config.samplingPercentage === 100;
  }

  /**
   * Returns the instrumentation key associated with the resource
   * @returns The usage data instrumentation key
   */
  public getUsageDataKey(): string {
    return this.options.instrumentationKey;
  }

  /**
   * Transform the project name by adddin '-test' suffix to it if necessary
   */
  private getEventName() {
    return this.options.isForTesting ? `${this.options.projectName}-test` : this.options.projectName;
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
   * Get the usage data level
   * @returns the usage data level
   */
  public getUsageDataLevel(): string {
    return this.options.usageDataLevel;
  }

  /**
   * Returns parsed file path, scrubbing file names and sensitive information
   * @returns Error after removing PII
   */
  public maskFilePaths(err: Error): Error {
    try {
      const regexRemoveUserFilePaths = /(\w:)*[/\\](.*[/\\]+)*(.+\.)+[a-zA-Z]+/gim;
      const maskToken: string = "<filepath>";

      err.message = err.message.replace(regexRemoveUserFilePaths, maskToken);
      err.stack = err.stack.replace(regexRemoveUserFilePaths, maskToken);

      return err;
    } catch (err) {
      this.reportError("maskFilePaths", err);
      throw new Error(err);
    }
  }

  /**
   * Removes sensitive information fields from ApplicationInsights data
   */
  private removeApplicationInsightsSensitiveInformation() {
    delete this.usageDataClient.context.tags["ai.cloud.roleInstance"]; // cloud name
    delete this.usageDataClient.context.tags["ai.device.id"]; // machine name
    delete this.usageDataClient.context.tags["ai.user.accountId"]; // subscription
  }

  /**
   * Reports custom exception event object to Application Insights
   * @param method Method name sent to Application Insights
   * @param err Error or message about error sent to Application Insights
   * @param data Data object(s) sent to Application Insights
   */
  public reportException(method: string, err: Error | string, data: object = {}) {
    if (this.getUsageDataLevel() === UsageDataLevel.on) {
      try {
        if (err instanceof ExpectedError) {
          this.reportExpectedException(method, err, data);
          return;
        }
        let error = err instanceof Error ? Object.create(err) : new Error(`${this.options.projectName} error: ${err}`);
        error.name = this.getEventName();
        let exceptionTelemetryObj: appInsights.Contracts.ExceptionTelemetry = {
          exception: this.maskFilePaths(error),
          properties: {},
        };
        Object.entries({
          Succeeded: false,
          Method: method,
          ExpectedError: false,
          ...this.defaultData,
          ...data,
          deviceID: this.options.deviceID,
        }).forEach((entry) => {
          exceptionTelemetryObj.properties[entry[0]] = JSON.stringify(entry[1]);
        });
        this.usageDataClient.trackException(exceptionTelemetryObj);
        this.exceptionsSent++;
      } catch (e) {
        this.reportError("reportException", e);
        throw e;
      }
    }
  }

  /**
   * Reports custom expected exception event object to Application Insights
   * @param method Method name sent to Application Insights
   * @param err Error or message about error sent to Application Insights
   * @param data Data object(s) sent to Application Insights
   */
  public reportExpectedException(method: string, err: Error | string, data: object = {}) {
    let error = err instanceof Error ? Object.create(err) : new Error(`${this.options.projectName} error: ${err}`);
    error.name = this.getEventName();
    this.maskFilePaths(error);
    const errorMessage = error instanceof Error ? error.message : error;

    this.sendUsageDataEvent({
      Succeeded: true,
      Method: method,
      ExpectedError: true,
      Error: errorMessage,
      ...data,
    });
  }

  /**
   * Reports custom success event object to Application Insights
   * @param method Method name sent to Application Insights
   * @param data Data object(s) sent to Application Insights
   */
  public reportSuccess(method: string, data: object = {}) {
    this.sendUsageDataEvent({
      Succeeded: true,
      Method: method,
      ExpectedError: false,
      ...data,
    });
  }

  /**
   * Reports custom exception event object to Application Insights
   * @param method Method name sent to Application Insights
   * @param err Error or message about error sent to Application Insights
   * @param data Data object(s) sent to Application Insights
   * @deprecated Use `reportUnexpectedError` instead.
   */
  public sendUsageDataException(method: string, err: Error | string, data: object = {}) {
    if (this.getUsageDataLevel() === UsageDataLevel.on) {
      try {
        let error = err instanceof Error ? Object.create(err) : new Error(`${this.options.projectName} error: ${err}`);
        error.name = this.getEventName();
        let exceptionTelemetryObj: appInsights.Contracts.ExceptionTelemetry = {
          exception: this.maskFilePaths(error),
          properties: {},
        };
        Object.entries({
          Succeeded: false,
          Method: method,
          ...this.defaultData,
          ...data,
        }).forEach((entry) => {
          exceptionTelemetryObj.properties[entry[0]] = JSON.stringify(entry[1]);
        });
        this.usageDataClient.trackException(exceptionTelemetryObj);
        this.exceptionsSent++;
      } catch (e) {
        this.reportError("sendUsageDataException", e);
        throw e;
      }
    }
  }

  /**
   * Reports custom success event object to Application Insights
   * @param method Method name sent to Application Insights
   * @param data Data object(s) sent to Application Insights
   * @deprecated Use `reportSuccess` instead.
   */
  public sendUsageDataSuccessEvent(method: string, data: object = {}) {
    this.sendUsageDataEvent({
      Succeeded: true,
      Method: method,
      Pass: true,
      ...data,
    });
  }

  /**
   * Reports custom successful fail event object to Application Insights
   * "Successful fail" means that there was an error as a result of user error, but our code worked properly
   * @param method Method name sent to Application Insights
   * @param data Data object(s) sent to Application Insights
   * @deprecated Use `reportExpectedError` instead.
   */
  public sendUsageDataSuccessfulFailEvent(method: string, data: object = {}) {
    this.sendUsageDataEvent({
      Succeeded: true,
      Method: method,
      Pass: false,
      ...data,
    });
  }

  /**
   * Reports custom event object to Application Insights
   * @param data Data object(s) sent to Application Insights
   */
  public sendUsageDataEvent(data: object = {}) {
    if (this.getUsageDataLevel() === UsageDataLevel.on) {
      try {
        let eventTelemetryObj = new appInsights.Contracts.EventData();
        eventTelemetryObj.name = this.getEventName();
        eventTelemetryObj.properties = {
          ...this.defaultData,
          ...data,
          deviceID: this.options.deviceID,
        };
        this.usageDataClient.trackEvent(eventTelemetryObj);
        this.eventsSent++;
      } catch (e) {
        this.reportError("sendUsageDataEvent", e);
      }
    }
  }
}
