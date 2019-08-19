import * as appInsights from "applicationinsights";
import * as readLine from "readline-sync";
import * as jsonData from "./usageDataJsonData";
/**
 * Specifies the usage-data infrastructure the user wishes to use
 * @enum Application Insights: Microsoft Azure service used to collect and query through data
 */
export enum UsageDataType {
  applicationinsights = "applicationInsights",
}
/**
 * Level controlling what type of usage-data is being sent
 * @enum off: off level of usage-data, sends no usage-data
 * @enum on: on level of usage-data, sends errors and events
 */
export enum UsageDataLevel {
  off = "off",
  on = "on",
}

/**
 * Usage-Data object necessary for initialization of usage-data package
 * @member groupName Usage-Data Group name that will be written to the usage-Data config file (i.e. usage-dataJsonFilePath)
 * @member projectName The name of the project that is using the usage-data package (e.g "generator-office")
 * @member instrumentationKey Instrumentation key for usage-data resource
 * @member promptQuestion Question displayed to user over opt-in for usage-data
 * @member raisePrompt Specifies whether to raise usage-data prompt (this allows for using a custom prompt)
 * @member usageDataLevel User's response to the prompt for usage-data
 * @member usageDataType Usage-Data infrastructure to send data
 * @member testData Allows user to run program without sending actual data
 */
export interface IUsageDataOptions {
  groupName: string;
  projectName: string;
  instrumentationKey: string;
  promptQuestion: string;
  raisePrompt: boolean;
  usageDataLevel: UsageDataLevel;
  usageDataType: UsageDataType;
  testData: boolean;
}

/**
 * Creates and initializes member variables while prompting user for usage-data collection when necessary
 * @param usageDataObject
 */
export class OfficeAddinUsageData {
  private usageDataClient = appInsights.defaultClient;
  private eventsSent: number = 0;
  private exceptionsSent: number = 0;
  private usageDataObject: IUsageDataOptions;

  constructor(usageDataOptions: IUsageDataOptions) {
    try {
      this.usageDataObject = usageDataOptions;

      if (this.usageDataObject.instrumentationKey === undefined) {
        throw new Error("Instrumentation Key not defined - cannot create usage-data object");
      }

      if (this.usageDataObject.groupName === undefined) {
        throw new Error("Group Name not defined - cannot create usage-data object");
      }

      if (this.usageDataObject.promptQuestion === undefined) {
        this.usageDataObject.promptQuestion = `Office Add-in CLI tools collect anonymized usage data which is sent to Microsoft to help improve our product. Please read our privacy statement and usage-data details at https://aka.ms/OfficeAddInCLIPrivacy.`;
      }

      if (jsonData.groupNameExists(this.usageDataObject.groupName)) {
        this.usageDataObject.usageDataLevel = jsonData.readUsageDataLevel(this.usageDataObject.groupName);
      }

      if (!this.usageDataObject.testData && this.usageDataObject.raisePrompt && jsonData.needToPromptForUsageData(this.usageDataObject.groupName)) {
        this.usageDataOptIn();
      } else {
        jsonData.writeUsageDataJsonData(this.usageDataObject.groupName, this.usageDataObject.usageDataLevel);
      }

      appInsights.setup(this.usageDataObject.instrumentationKey)
        .setAutoCollectExceptions(false)
        .start();
      this.usageDataClient = appInsights.defaultClient;
      this.removeApplicationInsightsSensitiveInformation();
    } catch (err) {
      throw new Error(err);
    }
  }

  /**
   * Reports custom event object to usage-data structure
   * @param eventName Event name sent to usage-data structure
   * @param data Data object sent to usage-data structure
   */
  public async reportEvent(eventName: string, data: object): Promise<void> {
    if (this.UsageDataLevel() === UsageDataLevel.on) {
      this.reportEventApplicationInsights(eventName, data);
    }
  }

  /**
   * Reports custom event object to Application Insights
   * @param eventName Event name sent to Application Insights
   * @param data Data object sent to Application Insights
   */
  public async reportEventApplicationInsights(eventName: string, data: object): Promise<void> {
    if (this.UsageDataLevel() === UsageDataLevel.on) {
      const usageDataEvent = new appInsights.Contracts.EventData();
      usageDataEvent.name = eventName;
      try {
        for (const [key, [value, elapsedTime]] of Object.entries(data)) {
          usageDataEvent.properties[key] = value;
          usageDataEvent.measurements[key + " durationElapsed"] = elapsedTime;
        }
        if (!this.usageDataObject.testData) {
          this.usageDataClient.trackEvent(usageDataEvent);
        }
        this.eventsSent++;
      } catch (err) {
        this.reportError("sendUsageDataEvents", err);
        throw new Error(err);
      }
    }
  }

  /**
   * Reports error to usage-data structure
   * @param errorName Error name sent to usage-data structure
   * @param err Error sent to usage-data structure
   */
  public async reportError(errorName: string, err: Error): Promise<void> {
    if (this.UsageDataLevel() === UsageDataLevel.on) {
      this.reportErrorApplicationInsights(errorName, err);
    }
  }

  /**
   * Reports error to Application Insights
   * @param errorName Error name sent to Application Insights
   * @param err Error sent to Application Insights
   */
  public async reportErrorApplicationInsights(errorName: string, err: Error): Promise<void> {
    if (this.UsageDataLevel() === UsageDataLevel.on) {
      err.name = errorName;
    if (!this.usageDataObject.testData) {
      this.usageDataClient.trackException({ exception: this.maskFilePaths(err) });
    }
    this.exceptionsSent++;
    }
  }
/**
 * Prompts user for usage-data participation once and records response
 * @param testData Specifies whether test code is calling this method
 * @param testReponse Specifies test response
 */
public usageDataOptIn(testData: boolean = this.usageDataObject.testData, testResponse: string = ""): void {
  try {
    let response: string = "";
    if (testData) {
      response = testResponse;
    } else {
      response = readLine.question(`${this.usageDataObject.promptQuestion}\n`);
    }
    if (response.toLowerCase() === "y") {
      this.usageDataObject.usageDataLevel = UsageDataLevel.on;
    } else {
      this.usageDataObject.usageDataLevel = UsageDataLevel.off;
    }
    jsonData.writeUsageDataJsonData(this.usageDataObject.groupName, this.usageDataObject.usageDataLevel);
  } catch (err) {
    this.reportError("UsageDataOptIn", err);
    throw new Error(err);
  }
 }
  /**
   * Stops usage-data from being sent, by default usage-data will be on
   */
  public setUsageDataOff() {
    appInsights.defaultClient.config.samplingPercentage = 0;
  }

  /**
   * Starts sending usage-data, by default usage-data will be on
   */
  public setUsageDataOn() {
    appInsights.defaultClient.config.samplingPercentage = 100;
  }

  /**
   * Returns whether the usage-data is currently on or off
   * @returns Whether usage-data is turned on or off
   */
  public isUsageDataOn(): boolean {
    return appInsights.defaultClient.config.samplingPercentage === 100;
  }

  /**
   * Returns the instrumentation key associated with the resource
   * @returns The usage-data instrumentation key
   */
  public getUsageDataKey(): string {
    return this.usageDataObject.instrumentationKey;
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
   * Returns whether usage-data is enabled on the current object
   * @returns Usage-Data is enabled or not
   */
  public UsageDataLevel(): string {
    return this.usageDataObject.usageDataLevel;
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
    delete this.usageDataClient.context.tags["ai.cloud.roleInstance"]; // cloud name
    delete this.usageDataClient.context.tags["ai.device.id"]; // machine name
    delete this.usageDataClient.context.tags["ai.user.accountId"]; // subscription
  }
}
