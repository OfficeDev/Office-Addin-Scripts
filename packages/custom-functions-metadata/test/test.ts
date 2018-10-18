import * as assert from "assert"; 
import * as mocha from "mocha"; 
import * as fs from "fs";
import * as ts from "typescript";
import * as jsongenerator from "../src/custom-functions-metadata";

describe("test json file created", function() {
    describe("generate test", function(){
        it("test it", function() {
            var inputFile = "../custom-functions-metadata/test/testfunctions.ts";
            var output = "./test.json";
            jsongenerator.generate(inputFile,output);
            jsongenerator.logError("testError");
            assert.strictEqual(jsongenerator.errorFound, true, "error not created");
            var skipped = 'notadded';
            assert.strictEqual(jsongenerator.skippedFunctions[0],skipped, "skipped function not found");
            assert.strictEqual(fs.existsSync(output), true, "json file not created");
        });
    });
});
describe("test errors", function() {
    describe("failure to generate", function(){
        it("test error", function() {
             var inputFile = "../custom-functions-metadata/test/errorfunctions.js";
            var output = "./errortest.json";
            jsongenerator.generate(inputFile,output);
            assert.strictEqual(jsongenerator.errorFound, true, "error not created");
            var errorlog = jsongenerator.errorLogFile[1];
            var errorstring = "Unsupported type in code comment:badtype";
            assert.strictEqual(errorlog, errorstring, "Error string not found.");
            assert.strictEqual(fs.existsSync(output), false, "json file created");
        });
    });
});