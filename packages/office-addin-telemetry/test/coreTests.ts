import * as assert from 'assert';
import * as fs from "fs";
import { OfficeAddinTelemetry, telemetryType } from "../src/officeAddinTelemetry";
import * as appInsights from "applicationinsights";
import * as path from 'path';
import * as mocha from "mocha";
<<<<<<< HEAD
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
=======
const addInTelemetry = new OfficeAddinTelemetry("de0d9e7c-1f46-4552-bc21-4e43e489a015", "", true);
var err = new Error("this error contains a file path: C://Users//t-juflor//AppData//Roaming//npm//node_modules//balanced-match//index.js");//for testing parse method

describe('reportEvent', () => {
  it('should track event of object passed in with a project name', () => {
    addInTelemetry.setTelemetryOff();
    var test1 = { "Test": true };
    addInTelemetry.reportEvent("TestData", test1);
    assert(1 === addInTelemetry.getEventsSent());
  });
});

describe('reportError', () => {
  it('should send telemetry execption', () => {
    addInTelemetry.setTelemetryOff();
    const exception = new Error("this error contains a file path: C://Users//t-juflor//AppData//Roaming//npm//node_modules//balanced-match//index.js");
    addInTelemetry.reportError("ReportErrorCheck", exception);
    assert(1 === addInTelemetry.getExceptionsSent());
  });
});

describe('addTelemetry', () => {
  it('should add object to telemetry', () => {
    var test = {};
    addInTelemetry.addTelemetry(test, "Test", true);
    assert(JSON.stringify(test) === JSON.stringify({ "Test": true }));
  });
});

describe('checkPrompt', () => {
  //const path = require('os').homedir()+ "/AppData/Local/Temp/check.txt";//specific to windows
  const path = require('os').homedir() + "/mochaTest"
  var mochaFileLocation = "/mochaTest";
  it('should check to see if it has writen to a file if not creates file and writes to it returns true', () => {
    if (fs.existsSync(path)) {
      fs.unlinkSync(path)//deletes file if it exists
    }
    assert(true === addInTelemetry.checkPrompt(mochaFileLocation));

  });
  it('should check to see if text is in file, if appropriate word is in, returns false', () => {
    if (fs.existsSync(path) && fs.readFileSync(path).includes("de0d9e7c-1f46-4552-bc21-4e43e489a015")) {
      assert(false === addInTelemetry.checkPrompt(mochaFileLocation));
    }
    if (fs.existsSync(path)) {
      fs.unlinkSync(path)//deletes file
    }

  });

  it('should check to see if text is in file if already created, if appropriate word is not in, returns true and writes to file', () => {

    fs.writeFileSync(path, "");

    assert(true === addInTelemetry.checkPrompt(mochaFileLocation));
    var text = fs.readFileSync(path, "utf8");
    if (text.includes('de0d9e7c-1f46-4552-bc21-4e43e489a015')) {
      var response = true;
    } else {
      response = false;
    }

    assert(true === response);

    if (fs.existsSync(path)) {
      fs.unlinkSync(path)//deletes file
    }

  });
});
describe('telemetryOptIn', () => {
  it('should display user asking to opt in, changes to true if user types y ', () => {
    addInTelemetry.telemetryOptIn(1);
    assert(true === addInTelemetry.telemetryOptedIn2());
  });
  it('should display user asking to opt in, changes to false if user types anything else then y ', () => {
    addInTelemetry.telemetryOptIn(2);
    assert(false === addInTelemetry.telemetryOptedIn2());
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
    var test1 = { "Test": true };
    addInTelemetry.reportEvent("TestData", test1);
    assert(1 === addInTelemetry.getEventsSent());
  });
});

describe('getExceptionsSent', () => {
  it('should return amount of exceptions successfully sent ', () => {
    addInTelemetry.setTelemetryOff();
    const exception = new Error("this error contains a file path: C://Users//t-juflor//AppData//Roaming//npm//node_modules//balanced-match//index.js");
    addInTelemetry.reportError("TestData", exception);
    assert(1 === addInTelemetry.getExceptionsSent());
  });
});

describe('telemetryOptedIn', () => {//could be connected with telemetryOptIn
  it('should return true if user opted in', () => {
    addInTelemetry.telemetryOptIn(1);
    assert(true === addInTelemetry.telemetryOptedIn2());

  });
  it('should return false if user opted out', () => {
    addInTelemetry.telemetryOptIn(2);
    assert(false === addInTelemetry.telemetryOptedIn2());
  });
});

describe('parseErrors', () => {
  it('should return a parsed file path error', () => {
    addInTelemetry.setTelemetryOff();
    err = addInTelemetry.mochaTestFilePaths(err);
    var compare_error = new Error();
    compare_error.message = "this error contains a file path: C:index.js";
    compare_error.stack = `Error: this error contains a file path: C:index.js
    at Object.<anonymous> (coreTests.ts:8:11)`;//may throw error if change any part of the top of the test file
    err = addInTelemetry.mochaTestFilePaths(err);
    //console.log(err.stack);
    //console.log(compare_error.stack);
    assert(compare_error.name === err.name);
    assert(compare_error.message === err.message);
    assert(err.stack.includes(compare_error.stack));
  });
});
>>>>>>> cd5d4caf695d7894ffbb275723679ba682ab49aa



