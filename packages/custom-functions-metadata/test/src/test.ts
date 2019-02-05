import * as assert from "assert";
import * as fs from "fs";
import * as mocha from "mocha";
import * as ts from "typescript";
import * as customFunctionsMetadata from "../../src/custom-functions-metadata";

describe("test json file created", function() {
    describe("generate test", function() {
        it("test it", async function() {
            const inputFile = "../custom-functions-metadata/test/typescript/testfunctions.ts";
            const output = "test.json";
            await customFunctionsMetadata.generate(inputFile, output, true);
            const skipped = "notAdded";
            assert.strictEqual(customFunctionsMetadata.skippedFunctions[0], skipped, "skipped function not found");
            assert.strictEqual(fs.existsSync(output), true, "json file not created");
        });
    });
});
describe("verify json created in file by typescript", function() {
    describe("verify metadata from typescript", function() {
        it("test json", function() {
            const output = "test.json";
            const jsonCreated = fs.readFileSync(output);
            const j = JSON.parse(jsonCreated.toString());
            assert.strictEqual(j.functions[0].id, "ADD", "id not created properly");
            assert.strictEqual(j.functions[0].name, "ADD", "name not created properly");
            assert.strictEqual(j.functions[0].description, "Test comments", "description not created properly");
            assert.strictEqual(j.functions[0].helpUrl, "https://dev.office.com", "helpUrl not created properly");
            assert.strictEqual(j.functions[0].parameters[0].name, "first", "parameter name not created properly");
            assert.strictEqual(j.functions[0].parameters[0].description, "the first number", "description not created properly");
            assert.strictEqual(j.functions[0].parameters[0].type, "number", "type not created properly");
            assert.strictEqual(j.functions[0].parameters[0].optional, false, "optional not created properly");
            assert.strictEqual(j.functions[0].result.type, "number", "result type not created properly");
            assert.strictEqual(j.functions[0].options.volatile, true, "options volatile not created properly");
            assert.strictEqual(j.functions[0].options.stream, true, "options stream not created properly");
            assert.strictEqual(j.functions[0].options.cancelable, true, "options cancelable not created properly");
            assert.strictEqual(j.functions[1].result.dimensionality, "matrix", "result dimensionality matrix not created properly");
            assert.strictEqual(j.functions[2].parameters[0].type, "boolean", "type boolean not created properly");
            assert.strictEqual(j.functions[2].result.type, "boolean", "result type boolean not created properly");
            assert.strictEqual(j.functions[4].parameters[0].type, "string", "type string not created properly");
            assert.strictEqual(j.functions[4].result.type, "string", "result type string not created properly");
            assert.strictEqual(j.functions[5].result.type, "any", "void type - result type any not created properly");
            assert.strictEqual(j.functions[6].parameters[0].type, "any", "object type - type any not created properly");
            assert.strictEqual(j.functions[6].result.type, "any", "object type - result type any not created properly");
            assert.strictEqual(j.functions[8].parameters[0].type, "any", "enum type - type any not created properly");
            assert.strictEqual(j.functions[8].parameters[0].dimensionality, "matrix", "enum type - parameter dimensionality matrix any not created properly");
            assert.strictEqual(j.functions[8].result.type, "any", "enum type - result type any not created properly");
            assert.strictEqual(j.functions[8].result.dimensionality, "matrix", "enum type - result dimensionality matrix any not created properly");
            assert.strictEqual(j.functions[9].parameters[0].type, "any", "tuple type - type any not created properly");
            assert.strictEqual(j.functions[9].result.type, "any", "tuple type - result type any not created properly");
            assert.strictEqual(j.functions[10].options.stream, true, "CustomFunctions.StreamingHandler - options stream not created properly");
            assert.strictEqual(j.functions[10].result.type, "number", "CustomFunctions.StreamingHandler - result type number not created properly");
            assert.strictEqual(j.functions[11].parameters[0].optional, true, "optional true not created properly");
            assert.strictEqual(j.functions[12].parameters[0].type, "any", "any type - type any not created properly");
            assert.strictEqual(j.functions[12].result.type, "any", "any type - result type any not created properly");
            assert.strictEqual(j.functions[13].options.cancelable, true, "CustomFunctions.CancelableHandler type not created properly");
            assert.strictEqual(j.functions[14].id, "UPDATEID", "@CustomFunction id not created properly");
            assert.strictEqual(j.functions[14].name, "updateName", "@CustomFunction name not created properly");
            assert.strictEqual(j.functions[15].options.requiresAddress, true, "requiresAddress tag not created properly");
            assert.strictEqual(j.functions[16].options.requiresAddress, true, "CustomFunctions.Invocation requiresAddress tag not created properly");
            assert.strictEqual(j.functions[17].options.cancelable, true, "CustomFunctions.CancelableInvocation type not created properly");
            assert.strictEqual(j.functions[18].options.stream, true, "CustomFunctions.StreamingInvocation - options stream not created properly");
        });
    });
});
describe("test javascript file as input", function() {
    describe("js test", function() {
        it("basic test", async function() {
            const inputFile = "../custom-functions-metadata/test/javascript/testjs.js";
            const output = "testjs.json";
            await customFunctionsMetadata.generate(inputFile, output);
            assert.strictEqual(fs.existsSync(output), true, "json file not created");
        });
    });
});
describe("verify json created in file by javascript", function() {
    describe("test javascript json", function() {
        it("test json", function() {
            const output = "testjs.json";
            const jsonCreated = fs.readFileSync(output);
            const j = JSON.parse(jsonCreated.toString());
            assert.strictEqual(j.functions[0].id, "TESTADD", "id not created properly");
            assert.strictEqual(j.functions[0].name, "TESTADD", "name not created properly");
            assert.strictEqual(j.functions[0].description, "This function is testing add", "description not created properly");
            assert.strictEqual(j.functions[0].parameters[0].name, "number1", "parameter name not created properly");
            assert.strictEqual(j.functions[0].parameters[0].description, "first number", "description not created properly");
            assert.strictEqual(j.functions[0].parameters[0].type, "number", "type not created properly");
            assert.strictEqual(j.functions[0].parameters[0].optional, false, "optional not created properly");
            assert.strictEqual(j.functions[0].result.type, "number", "result type not created properly");
            assert.strictEqual(j.functions[1].parameters[0].type, "boolean", "type boolean not created properly");
            assert.strictEqual(j.functions[1].result.type, "boolean", "result type boolean not created properly");
            assert.strictEqual(j.functions[2].parameters[0].optional, true, "optional true not created properly");
            assert.strictEqual(j.functions[3].parameters[0].type, "string", "type string not created properly");
            assert.strictEqual(j.functions[3].result.type, "string", "result type string not created properly");
            assert.strictEqual(j.functions[4].parameters[0].type, "any", "type any not created properly");
            assert.strictEqual(j.functions[4].result.type, "any", "result type any not created properly");
            assert.strictEqual(j.functions[5].options.stream, true, "CustomFunctions.StreamingHandler type any not created properly");
            assert.strictEqual(j.functions[5].result.type, "string", "streaming result type any not created properly");
            assert.strictEqual(j.functions[6].options.cancelable, true, "CustomFunctions.CancelableHandler type any not created properly");
            assert.strictEqual(j.functions[7].id, "NEWID", "@CustomFunction id not created properly");
            assert.strictEqual(j.functions[7].name, "NEWID", "@CustomFunction id set for name not created properly");
            assert.strictEqual(j.functions[8].id, "NEWID", "@CustomFunction id name not created properly");
            assert.strictEqual(j.functions[8].name, "newName", "@CustomFunction id name not created properly");
            assert.strictEqual(j.functions[9].options.requiresAddress, true, "CustomFunctions.Invocation set requiresAddress not created properly");
            assert.strictEqual(j.functions[10].options.stream, true, "CustomFunctions.StreamingInvocation type any not created properly");
            assert.strictEqual(j.functions[11].options.cancelable, true, "CustomFunctions.CancelableInvocation type any not created properly");
        });
    });
});
describe("test errors", function() {
    describe("failure to generate", function() {
        it("test error", async function() {
             const inputFile = "../custom-functions-metadata/test/javascript/errorfunctions.js";
             const output = "./errortest.json";
             await customFunctionsMetadata.generate(inputFile, output, true);
             const errtest: string[] = customFunctionsMetadata.errorLogFile;
             const errorIdBad = "ID-BAD";
             const errorNameBad = "1invalidname";
             const errorstring = "Unsupported type in code comment:badtype";
             assert.equal(errtest[0].includes(errorstring), true, "Unsupported type found");
             assert.equal(errtest[2].includes(errorIdBad), true, "Invalid id found");
             assert.equal(errtest[4].includes(errorNameBad), true, "Invalid name found");
             assert.strictEqual(fs.existsSync(output), false, "json file created");
        });
    });
});
describe("test bad file paths", function() {
    describe("failure to generate bad file path", function() {
        it("test error file path", async function() {
            const inputFile = "doesnotexist.ts";
            const output = "./nofile.json";
            const testError = "ENOENT: no such file or directory";
            try {
                await customFunctionsMetadata.generate(inputFile, output, true);
            } catch (error) {
                assert.ok(error.message.startsWith(testError), "Error message not found");
                assert.ok(error.message.includes(inputFile), "File name not found in error message");

            }
            assert.strictEqual(fs.existsSync(output), false, "json file created");
        });
    });
});
describe("delete test files", function() {
    describe("deleting files", function() {
        it("files to delete", function() {
            const outputJavaScript = "testjs.json";
            const outputTypeScript = "test.json";
            fs.unlinkSync(outputJavaScript);
            fs.unlinkSync(outputTypeScript);
        });
    });
});
