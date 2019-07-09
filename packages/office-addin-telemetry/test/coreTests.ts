import * as assert from 'assert';
import * as fs from "fs";
import { OfficeAddinTelemetry } from "../src/officeAddinTelemetry";
import * as appInsights from "applicationinsights";
import * as path from 'path';
import {  describe, before, it } from 'mocha';
const addInTelemetry = new OfficeAddinTelemetry("de0d9e7c-1f46-4552-bc21-4e43e489a015", "",true);
    
    describe('reportEvent', () => {
    it('should track event of object passed in with a project name', () => {
        addInTelemetry.setTelemetryOff();
        var test1 = {"Test":true};
        addInTelemetry.reportEvent("TestData",test1);
        assert(1 === addInTelemetry.getEventsSent());
    });
  });

    describe('reportError', () => {
    it('should send telemetry execption', () => {
        addInTelemetry.setTelemetryOff();
        const exception = new Error("this error contains a file path: C://Users//t-juflor//AppData//Roaming//npm//node_modules//balanced-match//index.js");
        addInTelemetry.reportError("ReportErrorCheck",exception);
        assert(1 === addInTelemetry.getExceptionsSent());
    });
  });

    describe('addTelemetry', () => {
        	it('should add object to telemetry', () => {
            	var test ={};
            	addInTelemetry.addTelemetry(test, "Test", true);
            	assert(JSON.stringify(test) === JSON.stringify({"Test": true }));
        	});
          });

    describe('checkPrompt', () => {


        	it('should check to see if it has writen to a file if not creates file and writes to it returns true', () => {
            if(fs.existsSync("./check.txt")){
            fs.unlinkSync('./check.txt')//deletes file
          }
          assert(true === addInTelemetry.checkPrompt());

          });

        
          it('should check to see if text is in file, if appropriate word is in, returns false', () => {

          assert(false === addInTelemetry.checkPrompt());

          if(fs.existsSync("./check.txt")){
            fs.unlinkSync('./check.txt')//deletes file
          }

          });

          it('should check to see if text is in file if already created, if appropriate word is not in, returns true and writes to file', () => {

          fs.writeFileSync("./check.txt", "");

          assert(true === addInTelemetry.checkPrompt());
          var text = fs.readFileSync("./check.txt","utf8");
          if (text === "done"){
            var response = true;
          }else{
            response = false;
          }

          assert(true === response);

          if(fs.existsSync("./check.txt")){
            fs.unlinkSync('./check.txt')//deletes file
          }

          });
        });
    describe('telemetryOptIn', () => {//TO DO
        	it('should display user asking to opt in, either changes m_telemetryOptIn to true or leaves it false ', () => {
                assert(true,)

        	});
          });


    describe('setTelemetryOff', () => {
        	it('should change samplingPercentage to 100, turns telemetry on', () => {
            	addInTelemetry.setTelemetryOn();
            	addInTelemetry.setTelemetryOff();
            	assert(0 === appInsights.defaultClient.config.samplingPercentage);
        	});
          });
          
    describe('setTelemetryOn', () => {
        	it('should change samplingPercentage to 100, turns telemetry on', () => {
            	addInTelemetry.setTelemetryOff();
            	addInTelemetry.setTelemetryOn();
            	assert(100 === appInsights.defaultClient.config.samplingPercentage);
        	});
          });
    describe('isTelemetryOn', () => {
        	it('should return true if samplingPercentage is on(100)', () => {
              appInsights.defaultClient.config.samplingPercentage = 100;
            	assert(true === addInTelemetry.isTelemetryOn());
          });
          
        	it('should return false if samplingPercentage is off(0)', () => {
              appInsights.defaultClient.config.samplingPercentage = 0;
            	assert(false === addInTelemetry.isTelemetryOn());
        	});
        });
        
 	describe('getTelemtryKey', () => {
        	it('should return telemetry key', () => {
            	assert('de0d9e7c-1f46-4552-bc21-4e43e489a015' === addInTelemetry.getTelemtryKey());
        	});
          });

    describe('getEventsSent', () => {
        	it('should return amount of events successfully sent', () => {
                addInTelemetry.setTelemetryOff();
                var test1 = {"Test":true};
                addInTelemetry.reportEvent("TestData",test1);
                assert(2 === addInTelemetry.getEventsSent());
        	});
          });

    describe('getExceptionsSent', () => {
        	it('should return amount of exceptions successfully sent ',() => {
                addInTelemetry.setTelemetryOff();
                const exception = new Error("this error contains a file path: C://Users//t-juflor//AppData//Roaming//npm//node_modules//balanced-match//index.js");
                addInTelemetry.reportError("TestData",exception);
                assert(2 === addInTelemetry.getExceptionsSent());
            });
          });

    describe('telemetryOptedIn', () => {//TO DO
        	it('should return true if user opted in', () => {
            assert("Telemetry will be sent!");
                
          });
          it('should return false if user opted out', () => {
            assert("You will not be sending telemetry");
        	});
          });

   describe('parseErrors', () => {//TO DO
            it('should return a parsed file path',() => {
                  addInTelemetry.setTelemetryOff();
                  var exceptionObject = {};
                  var err = new Error("this error contains a file path: C://Users//t-juflor//AppData//Roaming//npm//node_modules//balanced-match//index.js");
                  var compare = new Error('this error contains a file path: C:index.js');
                  compare.stack = "";
                  compare.message = "this error contains a file path: C:index.js"
                  this.addTelemetry(exceptionObject, "EventName", "Tester");
  	              this.addTelemetry(exceptionObject, "Message", err.message);
  	              this.addTelemetry(exceptionObject, "Stack", err.stack);
                  addInTelemetry.parseException2(exceptionObject)
                  console.log(JSON.stringify(exceptionObject));
                  assert(compare ===  addInTelemetry.parseException2(exceptionObject));
              });
            });



