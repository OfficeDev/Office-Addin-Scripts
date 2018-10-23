import * as assert from "assert"; 
import * as mocha from "mocha"; 
import * as fs from "fs";
import * as ts from "typescript";
import * as jsongenerator from "../src/custom-functions-metadata";

describe("test json file created", function() {
    describe("generate test", function(){
        it("test it", function() {
            var inputFile = "../custom-functions-metadata/test/testfunctions.ts";
            var output = "./test/test.json";
            jsongenerator.generate(inputFile,output);
            var skipped = 'notAdded';
            assert.strictEqual(jsongenerator.skippedFunctions[0],skipped, "skipped function not found");
            assert.strictEqual(fs.existsSync(output), true, "json file not created");
        });
    });
});
describe("verify json created in file by typescript", function() {
    describe("verify metadata from typescript", function(){
        it("test json", function() {
            var output = "./test/test.json";
            let jsonCreated = fs.readFileSync(output);
            var j = JSON.parse(jsonCreated.toString());
            assert.strictEqual(j.functions[0].id,"add", "id not created properly");
            assert.strictEqual(j.functions[0].name,"ADD", "name not created properly");
            assert.strictEqual(j.functions[0].description,"Test comments", "description not created properly");
            assert.strictEqual(j.functions[0].helpUrl,"https://dev.office.com", "helpUrl not created properly");
            assert.strictEqual(j.functions[0].parameters[0].name,"first", "parameter name not created properly");
            assert.strictEqual(j.functions[0].parameters[0].description,"the first number", "description not created properly");
            assert.strictEqual(j.functions[0].parameters[0].type,"number", "type not created properly");
            assert.strictEqual(j.functions[0].parameters[0].optional,false, "optional not created properly");
            assert.strictEqual(j.functions[0].result.type,"number", "result type not created properly");
            assert.strictEqual(j.functions[0].options.volatile,true, "options volatile not created properly");
            assert.strictEqual(j.functions[0].options.stream,true, "options stream not created properly");
            assert.strictEqual(j.functions[0].options.cancelable,true, "options cancelable not created properly");
        });
    });
});
describe("test javascript file as input", function() {
    describe("js test", function(){
        it("basic test", function() {
            var inputFile = "../custom-functions-metadata/test/testjs.js";
            var output = "./test/testjs.json";
            jsongenerator.generate(inputFile,output);
            assert.strictEqual(fs.existsSync(output), true, "json file not created");
        });
    });
});
describe("verify json created in file by javascript", function() {
    describe("test javascript json", function(){
        it("test json", function() {
            var output = "./test/testjs.json";
            let jsonCreated = fs.readFileSync(output);
            var j = JSON.parse(jsonCreated.toString());
            assert.strictEqual(j.functions[0].id,"testAdd", "id not created properly");
            assert.strictEqual(j.functions[0].name,"TESTADD", "name not created properly");
            assert.strictEqual(j.functions[0].description,"This function is testing add", "description not created properly");
            assert.strictEqual(j.functions[0].parameters[0].name,"number1", "parameter name not created properly");
            assert.strictEqual(j.functions[0].parameters[0].description,"first number", "description not created properly");
            assert.strictEqual(j.functions[0].parameters[0].type,"number", "type not created properly");
            assert.strictEqual(j.functions[0].parameters[0].optional,false, "optional not created properly");
            assert.strictEqual(j.functions[0].result.type,"number", "result type not created properly");
        });
    });
});
describe("test errors", function() {
    describe("failure to generate", function(){
        it("test error", function() {
             var inputFile = "../custom-functions-metadata/test/errorfunctions.js";
            var output = "./errortest.json";
            jsongenerator.generate(inputFile,output);
            var errorlog = jsongenerator.errorLogFile[0];
            var errorstring = "Unsupported type in code comment:badtype";
            assert.strictEqual(errorlog, errorstring, "Error string not found.");
            assert.strictEqual(fs.existsSync(output), false, "json file created");
        });
    });
});
describe("test bad file paths", function() {
    describe("failure to generate bad file path", function(){
        it("test error file path", function() {
            var inputFile = "doesnotexist.ts";
            var output = "./nofile.json";
            var testError = "ENOENT: no such file or directory"
            try {
                jsongenerator.generate(inputFile,output);
            }
            catch (error) {
                assert.ok(error.message.startsWith(testError), "Error message not found");
                assert.ok(error.message.includes(inputFile), "File name not found in error message");

            }
            assert.strictEqual(fs.existsSync(output), false, "json file created");
        });
    });
});

