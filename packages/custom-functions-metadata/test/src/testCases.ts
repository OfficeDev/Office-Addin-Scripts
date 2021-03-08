// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as assert from "assert";
import * as fs from "fs";
import * as mocha from "mocha";
import * as path from "path";
import { generateCustomFunctionsMetadata } from "../../src/generate";

function deleteFileIfExists(filePath: string): void {
    if (fs.existsSync(filePath)) {
        fs.unlinkSync(filePath);
    }
}

function normalizeLineEndings(text: string | undefined): string | undefined {
    return text ? text.replace(/\r\n|\r/g, "\n") : text;
}

function readFileIfExists(filePath: string): string | undefined {
    return fs.existsSync(filePath) ? normalizeLineEndings(fs.readFileSync(filePath).toString()) : undefined;
}

describe("test cases", function() {
    const testCasesDirPath = path.resolve("./test/cases");
    const testCases = fs.readdirSync(testCasesDirPath);

    testCases.forEach((testCaseDirName: string) => {
        ["ts", "js"].forEach((scriptType: string) => {
            const testCaseDirPath = path.resolve(testCasesDirPath, testCaseDirName);
            const sourceFileName = `functions.${scriptType}`;
            const sourceFile = path.resolve(testCaseDirPath, sourceFileName);
            const source: string | undefined = readFileIfExists(sourceFile);

            if (source) {
                it(`${testCaseDirName}\\${sourceFileName}`, async function() {
                    // add a file named "skip" to skip the test case
                    // add an expression in the file and it will be skipped if not true
                    const skip: string | undefined = readFileIfExists(path.resolve(testCaseDirPath, "skip"));
                    if (skip !== undefined) {
                        // tslint:disable-next-line: no-eval
                        const skipResult = eval(skip);
                        if (!skipResult) {
                            this.skip();
                        }
                    } else {
                        const actualErrorsFile = path.join(testCaseDirPath, `actual.${scriptType}.errors.txt`);
                        const expectedErrorsFile = path.join(testCaseDirPath, `expected.${scriptType}.errors.txt`);
                        const actualMetadataFile = path.join(testCaseDirPath, `actual.${scriptType}.json`);
                        const expectedMetadataFile = path.join(testCaseDirPath, "expected.json");
                        const expectedMetadata = readFileIfExists(expectedMetadataFile) || "";

                        // add a file named "debugger" to break on the test case
                        if (fs.existsSync(path.resolve(testCaseDirPath, "debugger"))) {
                            // tslint:disable-next-line: no-debugger
                            debugger;
                        }

                        // generate metadata
                        const result = await generateCustomFunctionsMetadata(sourceFile);

                        const actualMetadata = result.metadataJson;
                        const actualErrors = (result.errors.length > 0) ? result.errors.join("\n") : undefined;
                        const expectedErrors = readFileIfExists(expectedErrorsFile);

                        if (result.errors.length > 0) {
                            deleteFileIfExists(actualMetadataFile);
                        } else {
                            // write the actual metadata file
                            fs.writeFileSync(actualMetadataFile, actualMetadata);
                        }

                        // if actual errors are different than expected, write out the actual errors to a file
                        // otherwise, delete the actual errors file if it exists
                        if (actualErrors !== expectedErrors) {
                            fs.writeFileSync(actualErrorsFile, actualErrors);
                        } else {
                            deleteFileIfExists(actualErrorsFile);
                        }

                        assert.strictEqual(actualMetadata, expectedMetadata, "metadata does not match expected");
                        assert.strictEqual(actualErrors, expectedErrors, "errors do not match expected");

                        // if actual metadata is what was expected, delete the actual metadata file
                        if (actualMetadata === expectedMetadata) {
                            deleteFileIfExists(actualMetadataFile);
                        }
                    }
                });
            }
        });
    });
});
