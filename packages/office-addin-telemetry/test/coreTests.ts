import * as assert from 'assert';
import * as fs from "fs";
import { OfficeAddinTelemetry, telemetryType } from "../src/officeAddinTelemetry";
import * as appInsights from "applicationinsights";
import * as path from 'path';
import * as mocha from "mocha";
const addInTelemetry = new OfficeAddinTelemetry("de0d9e7c-1f46-4552-bc21-4e43e489a015", telemetryType.applicationinsights,true);
const err = new Error("this error contains a file path:C://" + require('os').homedir()+"/AppData//Roaming//npm//node_modules//balanced-match//index.js");//for testing parse method
const testDataFileLocation = "/mochaTest";

    describe('test reportEvent method', () => {
    it('should track event of object passed in with a project name', () => {
        addInTelemetry.setTelemetryOff();
        var test1 = {"Test":true};
        addInTelemetry.reportEvent("TestData",test1);
        assert(1 === addInTelemetry.getEventsSent());
    });
  });

    describe(' test reportError method', () => {
    it('should send telemetry execption', () => {
        addInTelemetry.setTelemetryOff();
        addInTelemetry.reportError("ReportErrorCheck",err);
        assert(1 === addInTelemetry.getExceptionsSent());
    });
  });

    describe(' test addTelemetry method', () => {
        	it('should add object to telemetry', () => {
            	var test ={};
            	addInTelemetry.addTelemetry(test, "Test", true);
            	assert(JSON.stringify(test) === JSON.stringify({"Test": true }));
        	});
          });

    describe('test checkPrompt method', () => {
          const path = require('os').homedir()+ testDataFileLocation;
          it('should check to see if it has writen to a file if not creates file and writes to it returns true', () => {
            if(fs.existsSync(path)){
            fs.unlinkSync(path)//deletes file if it exists
          }
          assert(true === addInTelemetry.checkPrompt(testDataFileLocation));

          });

          it('should check to see if text is in file if already created, if appropriate word is not in, returns true and writes to file', () => {
          if(!fs.existsSync(path)){
          fs.writeFileSync(path, "");
          }
          assert(true === addInTelemetry.checkPrompt(testDataFileLocation));
          var text = fs.readFileSync(path,"utf8");
          if (text.includes(addInTelemetry.getTelemtryKey())){
            var response = true;
          }else{
            response = false;
          }

          assert(true === response);

          if(fs.existsSync(path)){
            fs.unlinkSync(path)//deletes file
          }

          });

          it('should check to see if text is in file, if appropriate word is in, returns false', () => {
            if(fs.existsSync(path) && fs.readFileSync(path).includes(addInTelemetry.getTelemtryKey())){
            assert(false === addInTelemetry.checkPrompt(testDataFileLocation));
            }
            if(fs.existsSync(path)){
              fs.unlinkSync(path)//deletes file
            }
  
            });
        });
    describe('test telemetryOptIn method', () => {
        	it('should display user asking to opt in, changes to true if user types y ', () => {
                addInTelemetry.telemetryOptIn(1);
                assert(true === addInTelemetry.telemetryOptedIn2());
          });
          it('should display user asking to opt in, changes to false if user types anything else then y ', () => {
            addInTelemetry.telemetryOptIn(2);
            assert(false === addInTelemetry.telemetryOptedIn2());
      });
          });


    describe('test setTelemetryOff method', () => {
        	it('should change samplingPercentage to 100, turns telemetry on', () => {
            	addInTelemetry.setTelemetryOn();
            	addInTelemetry.setTelemetryOff();
            	assert(0 === appInsights.defaultClient.config.samplingPercentage);
        	});
          });
          
    describe('test setTelemetryOn method', () => {
        	it('should change samplingPercentage to 100, turns telemetry on', () => {
            	addInTelemetry.setTelemetryOff();
            	addInTelemetry.setTelemetryOn();
            	assert(100 === appInsights.defaultClient.config.samplingPercentage);
        	});
          });
    describe('test isTelemetryOn method', () => {
        	it('should return true if samplingPercentage is on(100)', () => {
              appInsights.defaultClient.config.samplingPercentage = 100;
            	assert(true === addInTelemetry.isTelemetryOn());
          });
          
        	it('should return false if samplingPercentage is off(0)', () => {
              appInsights.defaultClient.config.samplingPercentage = 0;
            	assert(false === addInTelemetry.isTelemetryOn());
        	});
        });
        
 	describe('test getTelemtryKey method', () => {
        	it('should return telemetry key', () => {
            	assert('de0d9e7c-1f46-4552-bc21-4e43e489a015' === addInTelemetry.getTelemtryKey());
        	});
          });

    describe('test getEventsSent method', () => {
        	it('should return amount of events successfully sent', () => {
                addInTelemetry.setTelemetryOff();
                var test1 = {"Test":true};
                addInTelemetry.reportEvent("TestData",test1);
                assert(1 === addInTelemetry.getEventsSent());
        	});
          });

    describe('test getExceptionsSent method', () => {
        	it('should return amount of exceptions successfully sent ',() => {
                addInTelemetry.setTelemetryOff();
                addInTelemetry.reportError("TestData",err);
                assert(1 === addInTelemetry.getExceptionsSent());
            });
          });

    describe('test telemetryOptedIn method', () => {//could be connected with telemetryOptIn
        	it('should return true if user opted in', () => {
            addInTelemetry.telemetryOptIn(1);
            assert(true === addInTelemetry.telemetryOptedIn2());
                
          });
          it('should return false if user opted out', () => {
            addInTelemetry.telemetryOptIn(2);
            assert(false === addInTelemetry.telemetryOptedIn2());
        	});
          });

   describe('test parseErrors method', () => {
            it('should return a parsed file path error',() => {
                  addInTelemetry.setTelemetryOff();
                  var err2 = addInTelemetry.testMaskFilePaths(err);
                  var compare_error =new Error();
                  compare_error.name = "ReportErrorCheck";
                  compare_error.message = "this error contains a file path:C:index.js";
                  compare_error.stack = `ReportErrorCheck: this error contains a file path:C:index.js
    at Object.<anonymous> (coreTests.ts:8:13)`;//may throw error if change any part of the top of the test file
    err2 = addInTelemetry.testMaskFilePaths(err);
                  assert(compare_error.name ===  err.name);
                  assert(compare_error.message ===  err.message);
                  assert(err.stack.includes(compare_error.stack));
              });
            });



