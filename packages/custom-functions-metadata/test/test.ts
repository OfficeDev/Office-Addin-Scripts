import * as assert from "assert"; 
import * as mocha from "mocha"; 
import * as fs from "fs";
import * as ts from "typescript";
import * as jsongenerator from "../src/custom-functions-metadata";
//import * as mini from "minimist";

var path = require("path");
var argv = require("optimist").demand('config').argv;
var configFilePath = argv.config;

assert.ok(fs.existsSync(configFilePath), 'config file not found:' + configFilePath);
var config = require('nconf').env().argv().file({file: configFilePath});

var goodString:string = "functions.json created for file: src/testfunctions.ts";

var pathKey = config.get('jsonfile');
var pathlocation = pathKey.path;

describe("test json file created", function() {
    describe("generate test", function(){
        it("test it", function() {
            assert.strictEqual("test","test");
            var inputFile = "../custom-functions-metadata/test/testfunctions.ts";
            var output = "./test.json";
            jsongenerator.generate(inputFile,output);
            jsongenerator.logError("testError");
            assert.strictEqual(jsongenerator.errorFound, true, "error not created");
            var skipped = 'notadded';
            assert.strictEqual(jsongenerator.skippedFunctions[0],skipped, "skipped function not found");
            //var e = jsongenerator.errorFound;
            //console.log(e);
            //console.log(jsongenerator.skippedFunctions)
            //console.log(fs.existsSync(output));
            //assert.equal(fs.existsSync(output), "json file not created");
            //assert.ok(fs.existsSync(output), "json file not created2");
            assert.strictEqual(fs.existsSync(output), true, "json file not created");
        });
    });
});