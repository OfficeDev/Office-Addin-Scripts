import * as appInsights from "applicationinsights";

import { Base, ExceptionData, ExceptionDetails, StackFrame } from 'applicationinsights/out/Declarations/Contracts';
import { inherits } from 'util';
import * as readline from "readline";//used
import * as fs from 'fs';//used
import * as chalk from 'chalk';//used
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

    if(telemetryTypes.toLowerCase() ==='applicationinsights'){//declaring telemetry structure
    this.m_telemetrySource = telemetryType.applicationinsights;
    }else if(telemetryTypes.toLowerCase() ==='oteljs'){
      this.m_telemetrySource = telemetryType.OtelJs;
    }


  	if(this.checkPrompt() || !mochaTest){
    	 this.telemetryOptIn();
    }

    	this.m_instrumentationKey = instrumentationKey;
    	appInsights.setup(this.m_instrumentationKey)
    	.setAutoCollectConsole(true)
    	.setAutoCollectDependencies(true)
    	.start()
      this.m_telemetryClient = appInsights.defaultClient;
      this.removeSensitiveInformation();
	}	

	public async reportEvent(eventName: string, data: object, elapsedTime = 0): Promise<void> {
  	if (this.telemetryOptedIn()) {
      	for (let [key, value] of Object.entries(data)) {
          	try {
                this.m_telemetryClient.trackEvent({ name: eventName, properties: { [key]: value }});
                this.events_sent++;
          	} catch (err) {
              	this.reportError("sendTelemetryEvents", err);
          	}
      	}
  	}
  }

	public async reportError(eventName: string, err: Error, mochaTest: boolean = false): Promise<void> {
    if (this.telemetryOptedIn()) {
  	const exceptionObject: object = {};
  	this.addTelemetry(exceptionObject, "EventName", eventName);
  	this.addTelemetry(exceptionObject, "Message", err.message);
  	this.addTelemetry(exceptionObject, "Stack", err.stack);
    this.m_telemetryClient.trackException({exception: this.parseException(exceptionObject)});
    if(mochaTest === true){
      
    }
    this.exceptions_sent++;
    }
  }
  
	public addTelemetry(data: {[k: string]: any}, key: string, value: any): object{
    	data[key] = value;
    	return data;
      	}


	public checkPrompt(): boolean{
    if(fs.existsSync("./check.txt")) {//checks to see if file exists
      	var text = fs.readFileSync("./check.txt","utf8");
     	if (text === "done"){
      	return false;
    	}else{
        fs.writeFileSync("./check.txt", "done");
        	return true;
      } 
    }else {
      	fs.writeFile("./check.txt", "done", (err) => {
        	if (err) console.log(err);

        	return true;
      	});
      }
    
    	return true;
  	}

	public telemetryOptIn(): void {
    var inquirer = require('inquirer');
    var chalk = require('chalk');
    var readlineSync = require('readline-sync');
    var response = readlineSync.question(chalk.blue('Do you want to opt in for telemetry [y/n]'));
    if(response.toLowerCase() === 'y'){
      this.m_telemetryOptIn = true;
      console.log('Telemetry will be sent!');
    }else {
      this.m_telemetryOptIn = false;
      console.log('You will not be sending telemetry');
    }
//
// get input from the user.
//
/*
inquirer
  .prompt([
    {
    message: 'What do you want to do?',
    }
  ])
  .then((answers:any) => {
    console.log(JSON.stringify(answers, null, '  '));
  });

/*
		const prompt = inquirer.prompt(chalk.cyan(`May this package report usage statistics to improve the tool over time?`));
    const { response } = await prompt;
    if(response.toLowerCase() === 'y'){
      this.m_telemetryOptIn = true;
      console.log('Telemetry will be sent!');
    }else{
      this.m_telemetryOptIn = false;
      console.log('You will not be sending telemetry');
    }*/
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
    delete this.m_telemetryClient.context.tags['ai.device.oemName'];//machine name
    delete this.m_telemetryClient.context.tags['ai.user.accountId'];//subscription #
  }


	private telemetryOptedIn(): boolean {
    	return this.m_telemetryOptIn;
	}
  public parseException2(err: Object): Error {
    return this.maskFilePaths(err);
  }
	private parseException(err: Object): Error {
  	return this.maskFilePaths(err);
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
            //console.log(prop + ': ' + stripedValue);
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
