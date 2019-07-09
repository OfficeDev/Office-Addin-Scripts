import * as appInsights from "applicationinsights";

import { Base, ExceptionData, ExceptionDetails, StackFrame } from 'applicationinsights/out/Declarations/Contracts';
import { inherits } from 'util';
import * as readline from "readline";//used
import * as fs from 'fs';//used
import * as chalk from 'chalk';//used
import * as os from "os";
export enum telemetryType {
  applicationinsights = "applicationinsights",
  OtelJs = "OtelJs",
}

export class OfficeAddinTelemetry {
	private m_instrumentationKey: string = "";
	private m_telemetryOptIn = true;//change to false once done
  private m_telemetryClient = appInsights.defaultClient;
  private m_appInsights: any;
  private m_telemetrySource = "";
  private events_sent = 0;
  private exceptions_sent = 0;


	constructor(instrumentationKey: string, telemetryTypes ="", mochaTest = false) {
    //checks to make sure it only displays the opt-in message once
    //telemetryType.applicationinsights = true;
    this.m_instrumentationKey = instrumentationKey;
    if(telemetryTypes.toLowerCase() ==='applicationinsights'){//declaring telemetry structure
    this.m_telemetrySource = telemetryType.applicationinsights;
    }else if(telemetryTypes.toLowerCase() ==='oteljs'){
      this.m_telemetrySource = telemetryType.OtelJs;
    }


  	if(this.checkPrompt() && !mochaTest){
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
      	for (let [key, value] of Object.entries(data)) {
          	try {
                //this.m_telemetryClient.trackEvent({ name: eventName, properties: { [key]: value }});
                this.events_sent++;
          	} catch (err) {
              	this.reportError("sendTelemetryEvents", err);
          	}
      	}
  	}
  }

	public async reportError(eventName: string, err: Error, mochaTest: boolean = false): Promise<void> {
    if (this.telemetryOptedIn()) {
    const regex = /\/(.*)\//gmi;
    err.message = err.message.replace(regex, "");
    //err.st
    err.stack = err.stackreplace(regex, "");
    // this.maskFilePaths(err);
    err.message = this.parseString(err.message);
    err.stack = this.parseString(err.stack);
  	const exceptionObject: object = {};
  	this.addTelemetry(exceptionObject, "EventName", eventName);
  	this.addTelemetry(exceptionObject, "Message", err.message);
  	this.addTelemetry(exceptionObject, "Stack", err.stack);
    //this.m_telemetryClient.trackException({exception: this.parseException(exceptionObject)});
    let exceptionObject2: object = {};
  	this.addTelemetry(exceptionObject2, "EventName", eventName);
  	this.addTelemetry(exceptionObject2, "Message", err.message);
    this.addTelemetry(exceptionObject2, "Stack", err.stack);
    exceptionObject2 = this.parseException(exceptionObject2);
    //console.log(JSON.stringify(exceptionObject2));
    this.exceptions_sent++;
    }
  }
  
	public addTelemetry(data: {[k: string]: any}, key: string, value: any): object{
    	data[key] = value;
    	return data;
      	}


	public checkPrompt(): boolean{
    const path = require('os').homedir()+ "/AppData/Local/Temp/check.txt";
    if(fs.existsSync(path)) {//checks to see if file exists
      	var text = fs.readFileSync(path,"utf8");
     	if (text.includes(this.getTelemtryKey())){
      	return false;
    	}else{
        fs.writeFileSync(path, this.getTelemtryKey());
        	return true;
      } 
    }else {
      	fs.writeFile(path, this.getTelemtryKey(), (err) => {
        	if (err) console.log(err);

        	return true;
      	});
      }
    
    	return true;
  	}

	public telemetryOptIn(mochaTest = 0): void {
    var chalk = require('chalk');
    var readlineSync = require('readline-sync');
    var response = "";
    if(mochaTest === 0){
    response = readlineSync.question(chalk.blue('Do you want to opt in for telemetry [y/n]'));
    }
    if(response.toLowerCase() === 'y' || mochaTest === 1){
      this.m_telemetryOptIn = true;
      console.log('Telemetry will be sent!');
    }else {
      this.m_telemetryOptIn = false;
      console.log('You will not be sending telemetry');
    }
	}

	public setTelemetryOff(){
  	appInsights.defaultClient.config.samplingPercentage = 0;
	}
	public setTelemetryOn(){
  	appInsights.defaultClient.config.samplingPercentage = 100;
	}
	public isTelemetryOn(): boolean{
  	if(appInsights.defaultClient.config.samplingPercentage === 100){
    	return true;
  	}else{
    	return false;
  	}
	}
	public getTelemtryKey(): string{
  	return this.m_instrumentationKey;
  }
  public getEventsSent(): any{
  	return this.events_sent;
  }
  public getExceptionsSent(): any{
  	return this.exceptions_sent;
  }
  
  private removeSensitiveInformation(){
    delete this.m_telemetryClient.context.tags['ai.cloud.roleInstance'];//cloud name
    delete this.m_telemetryClient.context.tags['ai.location.ip'];//location
    delete this.m_telemetryClient.context.tags['ai.device.id'];//machine name
    delete this.m_telemetryClient.context.tags['ai.user.accountId'];//subscription #
  }

  public telemetryOptedIn2(): boolean {//for mocha test
    return this.m_telemetryOptIn;
}

	private telemetryOptedIn(): boolean {
    	return this.m_telemetryOptIn;
	}
  public parseException2(err: Object): Error {//for mocha test
    return this.maskFilePaths(err);
  }
	private parseException(err: Object): Error {
  	return this.maskFilePaths(err);
  }
  private parseString(stack: any): string {
  	return this.maskFilePaths(stack);
  }

  private maskFilePaths(objString: any): any {
  	let obj: any = objString;
  	let isString: boolean = false;
  	if (typeof objString === "string") {
       	obj = JSON.parse(objString);
      	isString = true;
  	}
    const regex = /\/(.*)\//gmi;
  	for (var prop in obj) {
      	if (obj.hasOwnProperty(prop)) {
      	let value = obj[prop]
      	if (typeof value === "string") {
            let stripedValue = value.replace(regex, ' ');
            console.log(stripedValue);
            obj[prop] = stripedValue;
      	}
      	}
    }
  	if (isString){
	
      	return JSON.stringify(obj);
  	}else{

      	return objString;
  	}
}

}
