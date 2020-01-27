// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as appInsightsWeb from '@microsoft/applicationinsights-web';
import * as readLine from "readline-sync";
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
 * UpdateData options
 * @member groupName Group name for usage data settings
 * @member projectName The name of the project that is using the usage data package.
 * @member instrumentationKey Instrumentation key for usage data resource
 * @member promptQuestion Question displayed to user over opt-in for usage data
 * @member raisePrompt Specifies whether to raise usage data prompt (this allows for using a custom prompt)
 * @member usageDataLevel User's response to the prompt for usage data
 * @member method The desired method to use for reporting usage data.
 * @member isForTesting True if the data is just for testing, false for actual data that should be reported.
 */
export interface IUsageDataOptions {
  groupName: string;
  projectName: string;
  instrumentationKey: string;
  promptQuestion: string;
  raisePrompt: boolean;
  usageDataLevel: UsageDataLevel;
  method: UsageDataReportingMethod;
  isForTesting: boolean;
}

/**
 * Creates and initializes member variables while prompting user for usage data collection when necessary
 * @param usageDataObject
 */
export class OfficeAddinUsageData {
  private appInsightsWeb: appInsightsWeb.ApplicationInsights;
  private eventsSent: number = 0;
  private exceptionsSent: number = 0;
  private options: IUsageDataOptions;

  constructor(usageDataOptions: IUsageDataOptions) {
    try {
      this.options = usageDataOptions;

      if (this.options.instrumentationKey === undefined) {
        throw new Error("Instrumentation Key not defined - cannot create usage data object");
      }

      if (this.options.groupName === undefined) {
        throw new Error("Group Name not defined - cannot create usage data object");
      }

      this.appInsightsWeb = new appInsightsWeb.ApplicationInsights({ config: {
        instrumentationKey: this.options.instrumentationKey,
        disableExceptionTracking: true
      }});
      this.appInsightsWeb.loadAppInsights();
      this.removeApplicationInsightsSensitiveInformation();

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
      try {
        this.appInsightsWeb.trackEvent({
          name: this.options.isForTesting ? `${eventName}-test` : eventName, 
          properties: data
        });
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
   * @param isWeb Error sent as part of ApplicationInsights-JS (if true) or as part of ApplicationInsights-node.js (if false)
   */
  public async reportErrorApplicationInsights(errorName: string, err: Error): Promise<void> {
    if (this.getUsageDataLevel() === UsageDataLevel.on) {
      err.name = this.options.isForTesting ? `${errorName}-test` : errorName;
      this.appInsightsWeb.trackException(new appInsightsWeb.Exception(this.appInsightsWeb.core.logger, this.maskFilePaths(err)));
      this.exceptionsSent++;
    }
  }
/**
 * Prompts user for usage data participation once and records response
 * @param testData Specifies whether test code is calling this method
 * @param testReponse Specifies test response
 */

  /**
   * Stops usage data from being sent, by default usage data will be on
   */
  public setUsageDataOff() {
    this.appInsightsWeb.config.samplingPercentage = 0;
  }

  /**
   * Starts sending usage data, by default usage data will be on and at 100 percent
   * @param percent The percentage of usage data being sent up. Optional.
   */
  public setUsageDataOn(percent: number = 100) {
    if (percent <= 0) {
      throw new RangeError('Please choose a number higher than 0 for the percentage');
    } else if (percent > 100) {
      throw new RangeError('Please choose 100 or lower for the percentage');
    }
    this.appInsightsWeb.config.samplingPercentage = percent;
  }

  /**
   * Returns whether the usage data is currently on or off
   * @returns Whether usage data is turned on or off
   */
  public isUsageDataOn(): boolean {
    return this.appInsightsWeb.config.samplingPercentage > 0;
  }

  /**
   * Returns whether the usage data is currently on or off
   * @returns Whether usage data is turned on or off
   */
  public getUsageDataPercentage(): number {
    return this.appInsightsWeb.config.samplingPercentage;
  }

  /**
   * Returns the instrumentation key associated with the resource
   * @returns The usage data instrumentation key
   */
  public getUsageDataKey(): string {
    return this.options.instrumentationKey;
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
      const regexRemoveUserFilePaths = /\/(.*)\//gmi;
      const regexRemoveUserFilePathsFromStack = /\w:\\(?:[^\\\s]+\\)+/gmi;
      err.message = err.message.replace(regexRemoveUserFilePaths, "");
      err.stack = err.stack.replace(regexRemoveUserFilePaths, "");
      err.stack = err.stack.replace(regexRemoveUserFilePathsFromStack, "");
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
    // TODO: figure out the best way to delete this data
    // delete this.appInsightsWeb.context.device.id;
    // delete this.appInsightsWeb.context.user.accountId;
    // delete this.appInsightsWeb.context.location.ip;
    // delete this.appInsightsWeb.context.user.config.accountId;
  }
}
