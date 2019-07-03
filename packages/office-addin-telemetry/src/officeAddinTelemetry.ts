import * as appInsights from "applicationinsights";

import { Base, ExceptionData, ExceptionDetails, StackFrame } from 'applicationinsights/out/Declarations/Contracts';
import { inherits } from 'util';
import * as readline from "readline";
import * as fs from 'fs';
export class OfficeAddinTelemetry {
    private m_instrumentationKey: string = "";
    private m_telemetryOptIn = true;
    private m_telemetryClient = appInsights.defaultClient;

    constructor(instrumentationKey: string) {
      //checks to make sure it only displays the opt-in message once
      if(this.checkPrompt()){
        this.telemetryOptIn();
      }
        this.m_instrumentationKey = instrumentationKey;
        appInsights.setup(this.m_instrumentationKey)
        .setAutoCollectConsole(true)
        .setAutoCollectDependencies(true)
        .start()
        
        this.m_telemetryClient = appInsights.defaultClient;
    }

    

    public async sendTelemetryEvents(projectName: string, data: object, elapsedTime = 0): Promise<void> {
      if (this.telemetryOptedIn()) {
          for (let [key, value] of Object.entries(data)) {
              try {
                  this.m_telemetryClient.trackEvent({ name: projectName, properties: { [key]: value }});
              } catch (err) {
                  this.sendTelemetryExecption(projectName, err);
              }
          }
      }
  }

    public async sendTelemetryExecption(projectName: string, err: any): Promise<void> {
      const exceptionObject: object = {};
      this.addTelemetry(exceptionObject, "ProjectName", projectName);
      this.addTelemetry(exceptionObject, "Message", err.message);
      this.addTelemetry(exceptionObject, "Stack", err.stack);
      const parsedException: any = this.parseException(exceptionObject);
      this.m_telemetryClient.trackException({exception: parsedException});
    
    }
    public addTelemetry(data: {[k: string]: any}, key: string, value: any): object{
        data[key] = value;
        return data;
          }


    public checkPrompt(): boolean{
      fs.exists("./check.txt", (exist) => {//checks to see if file exists
        if (exist) {
          var text = fs.readFileSync("./check.txt","utf8");
         if (text === "done"){
          return false;
        }else{
          fs.writeFile("./check.txt", "done", (err) => {
            if (err) console.log(err);

            return true;
          });
        }
        } else {
          fs.writeFile("./check.txt", "done", (err) => {
            if (err) console.log(err);

            return false;
          });
        }
      });
        return true;
      }


    public telemetryOptIn(): void {
      let rl = readline.createInterface({
        input: process.stdin,
        output: process.stdout
      });
      
      rl.question('Do you want to opt in for telemetry [y/n] ', (answer) => {
        switch(answer.toLowerCase()) {
          case 'y':
            console.log('Super!');
            break;
          case 'n':
            console.log('Sorry! :(');
            break;
          default:
            console.log('Invalid answer!');
        }
        rl.close();
      });
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

    private telemetryOptedIn(): boolean {
        return this.m_telemetryOptIn;
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
              console.log(prop + ': ' + stripedValue);
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