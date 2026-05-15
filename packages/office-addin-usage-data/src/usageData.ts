// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

/* eslint-disable @typescript-eslint/no-unused-vars */

/**
 * Specifies the usage data infrastructure the user wishes to use
 * @enum Application Insights: Microsoft Azure service used to collect and query through data
 * @deprecated Usage data reporting has been removed
 */
export enum UsageDataReportingMethod {
  applicationInsights = "applicationInsights",
}
/**
 * Level controlling what type of usage data is being sent
 * @enum off: off level of usage data, sends no usage data
 * @enum on: on level of usage data, sends errors and events
 * @deprecated Usage data reporting has been removed
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
 * @deprecated Usage data reporting has been removed
 */
export interface IUsageDataOptions {
  groupName?: string;
  projectName: string;
  connectionString?: string;
  instrumentationKey?: string;
  promptQuestion?: string;
  raisePrompt?: boolean;
  usageDataLevel?: UsageDataLevel;
  method?: UsageDataReportingMethod;
  isForTesting?: boolean;
  deviceID?: string;
}

/**
 * Usage data class - all methods are now no-ops since telemetry has been removed
 * @deprecated Usage data reporting has been removed. This class is retained for API compatibility only.
 */
export class OfficeAddinUsageData {
  private options: IUsageDataOptions;

  constructor(usageDataOptions: IUsageDataOptions) {
    this.options = {
      usageDataLevel: UsageDataLevel.off,
      ...usageDataOptions,
    };
  }

  /** @deprecated No-op - usage data reporting has been removed */
  public async reportEvent(_eventName: string, _data: object): Promise<void> {}

  /** @deprecated No-op - usage data reporting has been removed */
  public async reportEventApplicationInsights(_eventName: string, _data: object): Promise<void> {}

  /** @deprecated No-op - usage data reporting has been removed */
  public async reportError(_errorName: string, _err: Error): Promise<void> {}

  /** @deprecated No-op - usage data reporting has been removed */
  public async reportErrorApplicationInsights(_errorName: string, _err: Error): Promise<void> {}

  /** @deprecated No-op - usage data reporting has been removed */
  public usageDataOptIn(_testData?: boolean, _testResponse?: string): void {}

  /** @deprecated No-op - usage data reporting has been removed */
  public setUsageDataOff(): void {}

  /** @deprecated No-op - usage data reporting has been removed */
  public setUsageDataOn(): void {}

  /** @deprecated Always returns false - usage data reporting has been removed */
  public isUsageDataOn(): boolean {
    return false;
  }

  /** @deprecated Returns empty string - usage data reporting has been removed */
  public getUsageDataKey(): string {
    return "";
  }

  /** @deprecated Always returns 0 - usage data reporting has been removed */
  public getEventsSent(): number {
    return 0;
  }

  /** @deprecated Always returns 0 - usage data reporting has been removed */
  public getExceptionsSent(): number {
    return 0;
  }

  /** @deprecated Always returns "off" - usage data reporting has been removed */
  public getUsageDataLevel(): string {
    return UsageDataLevel.off;
  }

  /** @deprecated Returns error unchanged - usage data reporting has been removed */
  public maskFilePaths(err: Error): Error {
    return err;
  }

  /** @deprecated No-op - usage data reporting has been removed */
  public reportException(_method: string, _err: Error | string, _data?: object): void {}

  /** @deprecated No-op - usage data reporting has been removed */
  public reportExpectedException(_method: string, _err: Error | string, _data?: object): void {}

  /** @deprecated No-op - usage data reporting has been removed */
  public reportSuccess(_method: string, _data?: object): void {}

  /** @deprecated No-op - usage data reporting has been removed */
  public sendUsageDataException(_method: string, _err: Error | string, _data?: object): void {}

  /** @deprecated No-op - usage data reporting has been removed */
  public sendUsageDataSuccessEvent(_method: string, _data?: object): void {}

  /** @deprecated No-op - usage data reporting has been removed */
  public sendUsageDataSuccessfulFailEvent(_method: string, _data?: object): void {}

  /** @deprecated No-op - usage data reporting has been removed */
  public sendUsageDataEvent(_data?: object): void {}
}
