// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as assert from "assert";
import { readFileSync } from "fs";
import * as mocha from "mocha";
import { generateCustomFunctionsMetadata } from "../../src/generate";
import { IParseTreeResult, parseTree } from "../../src/parseTree";

describe("test json output", function() {
    describe("generate test", function() {
        it("test it", async function() {
            const inputFile = "./test/typescript/testfunctions.ts";
            const output = "test.json";
            const generateResult = await generateCustomFunctionsMetadata(inputFile);
            const j = JSON.parse(generateResult.metadataJson);

            assert.strictEqual(generateResult.associate.length, 20, "associate array not complete");
            assert.strictEqual(generateResult.associate[0].id, "ADD", "associate id not created");
            assert.strictEqual(generateResult.associate[0].functionName, "add", "associate function name not created");
            assert.strictEqual(j.functions[0].id, "ADD", "id not created properly");
            assert.strictEqual(j.functions[0].name, "ADD", "name not created properly");
            assert.strictEqual(j.functions[0].description, "Test comments", "description not created properly");
            assert.strictEqual(j.functions[0].helpUrl, "https://docs.microsoft.com/office/dev/add-ins", "helpUrl not created properly");
            assert.strictEqual(j.functions[0].parameters[0].name, "first", "parameter name not created properly");
            assert.strictEqual(j.functions[0].parameters[0].description, "the first number", "description not created properly");
            assert.strictEqual(j.functions[0].parameters[0].type, "number", "type not created properly");
            assert.strictEqual(j.functions[0].parameters[0].optional, undefined, "optional not created properly");
            assert.strictEqual(j.functions[0].result.type, "number", "result type not created properly");
            assert.strictEqual(j.functions[0].options.volatile, true, "options volatile not created properly");
            assert.strictEqual(j.functions[0].options.stream, true, "options stream not created properly");
            assert.strictEqual(j.functions[0].options.cancelable, true, "options cancelable not created properly");
            assert.strictEqual(j.functions[1].result.dimensionality, "matrix", "result dimensionality matrix not created properly");
            assert.strictEqual(j.functions[2].parameters[0].type, "boolean", "type boolean not created properly");
            assert.strictEqual(j.functions[2].result.type, "boolean", "result type boolean not created properly");
            assert.strictEqual(j.functions[4].parameters[0].type, "string", "type string not created properly");
            assert.strictEqual(j.functions[4].result.type, "string", "result type string not created properly");
            assert.strictEqual(j.functions[5].result.type, undefined, "void type - result type any not created properly");
            assert.strictEqual(j.functions[6].parameters[0].type, "any", "object type - type any not created properly");
            assert.strictEqual(j.functions[6].result.type, undefined, "object type - result type any not created properly");
            assert.strictEqual(j.functions[8].parameters[0].type, "any", "enum type - type any not created properly");
            assert.strictEqual(j.functions[8].result.type, undefined, "enum type - result type any not created properly");
            assert.strictEqual(j.functions[9].parameters[0].type, "any", "tuple type - type any not created properly");
            assert.strictEqual(j.functions[9].result.type, undefined, "tuple type - result type any not created properly");
            assert.strictEqual(j.functions[10].options.stream, true, "CustomFunctions.StreamingHandler - options stream not created properly");
            assert.strictEqual(j.functions[10].result.type, "number", "CustomFunctions.StreamingHandler - result type number not created properly");
            assert.strictEqual(j.functions[11].parameters[0].optional, true, "optional true not created properly");
            assert.strictEqual(j.functions[12].parameters[0].type, "any", "any type - type any not created properly");
            assert.strictEqual(j.functions[12].result.type, undefined, "any type - result type any not created properly");
            assert.strictEqual(j.functions[13].options.cancelable, true, "CustomFunctions.CancelableHandler type not created properly");
            assert.strictEqual(j.functions[14].id, "UPDATEID", "@CustomFunction id not created properly");
            assert.strictEqual(j.functions[14].name, "updateName", "@CustomFunction name not created properly");
            assert.strictEqual(j.functions[15].options.requiresAddress, true, "requiresAddress tag not created properly");
            assert.strictEqual(j.functions[16].options.requiresAddress, true, "CustomFunctions.Invocation requiresAddress tag not created properly");
            assert.strictEqual(j.functions[17].options.cancelable, true, "CustomFunctions.CancelableInvocation type not created properly");
            assert.strictEqual(j.functions[17].options.requiresAddress, true, "CustomFunctions.CancelableInvocation requiresAdress type not created properly");
            assert.strictEqual(j.functions[18].options.stream, true, "CustomFunctions.StreamingInvocation - options stream not created properly");
            assert.strictEqual(j.functions[18].options.requiresAddress, undefined, "CustomFunctions.StreamingInvocation requiresAddress - should not be present");
            assert.strictEqual(j.functions[19].name, "UPPERCASE", "uppercased function name not created properly");
        });
    });
});
describe("test javascript file as input", function() {
    describe("js test", function() {
        it("basic test", async function() {
            const inputFile = "./test/javascript/testjs.js";
            const results = await generateCustomFunctionsMetadata(inputFile, true);
            const j = JSON.parse(results.metadataJson);

            assert.strictEqual(j.functions[0].id, "TESTADD", "id not created properly");
            assert.strictEqual(j.functions[0].name, "TESTADD", "name not created properly");
            assert.strictEqual(j.functions[0].description, "This function is testing add", "description not created properly");
            assert.strictEqual(j.functions[0].parameters[0].name, "number1", "parameter name not created properly");
            assert.strictEqual(j.functions[0].parameters[0].description, "first number", "description not created properly");
            assert.strictEqual(j.functions[0].parameters[0].type, "number", "type not created properly");
            assert.strictEqual(j.functions[0].parameters[0].optional, undefined, "optional not created properly");
            assert.strictEqual(j.functions[0].result.type, "number", "result type not created properly");
            assert.strictEqual(j.functions[1].parameters[0].type, "boolean", "type boolean not created properly");
            assert.strictEqual(j.functions[1].result.type, "boolean", "result type boolean not created properly");
            assert.strictEqual(j.functions[2].parameters[0].optional, true, "optional true not created properly");
            assert.strictEqual(j.functions[3].parameters[0].type, "string", "type string not created properly");
            assert.strictEqual(j.functions[3].result.type, "string", "result type string not created properly");
            assert.strictEqual(j.functions[4].parameters[0].type, "any", "type any not created properly");
            assert.strictEqual(j.functions[4].result.type, undefined, "result type any not created properly");
            assert.strictEqual(j.functions[5].options.stream, true, "CustomFunctions.StreamingHandler type any not created properly");
            assert.strictEqual(j.functions[5].result.type, "string", "streaming result type any not created properly");
            assert.strictEqual(j.functions[6].options.cancelable, true, "CustomFunctions.CancelableHandler type any not created properly");
            assert.strictEqual(j.functions[7].id, "NEWIDTEST", "@CustomFunction id not created properly");
            assert.strictEqual(j.functions[7].name, "NEWIDTEST", "@CustomFunction id set for name not created properly");
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
             const inputFile = "./test/javascript/errorfunctions.js";
             const generateResult = await generateCustomFunctionsMetadata(inputFile);
             const errtest: string[] = generateResult.errors;
             const errorIdBad = "ID-BAD";
             const errorNameBad = "1invalidname";
             const errorstring = "Custom function does not support type \"badtype\" as input or return parameter.";
             const errorPosition = "(7,12)";
             const errorRequiresAddress = "@requiresAddress";
             assert.strictEqual(errtest[0].includes(errorstring), true, "Unsupported type found");
             assert.strictEqual(errtest[0].includes(errorPosition), true, "Line and column number found");
             assert.strictEqual(errtest[2].includes(errorIdBad), true, "Invalid id found");
             assert.strictEqual(errtest[4], `The custom function name "1invalidname" should start with an alphabetic character. (25,19)`);
             assert.strictEqual(errtest[5], `The custom function name "1invalidname" should contain only alphabetic characters, numbers (0-9), period (.), and underscore (_). (25,19)`);
             assert.strictEqual(generateResult.metadataJson, "", "should not be any metadata");
        });
    });
});
describe("test parseTreeResult", function() {
    describe("parseTreeResult", function() {
        it("parseTree for errorfunctions", async function() {
            const inputFile = "./test/javascript/errorfunctions.js";
            const sourceCode = readFileSync(inputFile, "utf-8");
            const parseTreeResult: IParseTreeResult = parseTree(sourceCode, "errorfunctions");
            assert.strictEqual(parseTreeResult.extras[0].javascriptFunctionName, "testadd", "Function testadd found");
            assert.strictEqual(parseTreeResult.extras[0].errors.length, 1, "Correct number of errors found(1)");
            assert.strictEqual(parseTreeResult.extras[2].javascriptFunctionName, "badId", "Function badId found");
            assert.strictEqual(parseTreeResult.extras[2].errors.length, 2, "Correct number of errors found(2)");
            assert.strictEqual(parseTreeResult.extras[5].javascriptFunctionName, "привет", "Function привет found");
            assert.strictEqual(parseTreeResult.extras[5].errors[0].includes("привет".toLocaleUpperCase()), true, "Error message contains function name");
            assert.strictEqual(parseTreeResult.extras[6].errors[0].includes("Duplicate function name"), true, "Error message contains duplicate function name");
            assert.strictEqual(parseTreeResult.extras[7].errors[0].includes("@customfunction tag specifies a duplicate name"), true, "Error message contains duplicate function name from custom function");
            assert.strictEqual(parseTreeResult.extras[9].errors[0].includes("@customfunction tag specifies a duplicate name"), true, "Duplicate name from custom function tags");
            assert.strictEqual(parseTreeResult.extras[10].errors[0].includes("@customfunction tag specifies a duplicate id"), true, "Duplicate id from custom function tags");
            assert.strictEqual(parseTreeResult.extras[11].errors[0].includes("@customfunction tag specifies a duplicate id"), true, "Duplicate id with function name from custom function tags");
        });
    });
});
