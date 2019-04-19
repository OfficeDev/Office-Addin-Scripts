import * as assert from "assert";
import * as fs from "fs";
import * as mocha from "mocha";
import * as path from "path";
import { generate } from "../../src/generate";

function deleteFileIfExists(filePath: string): void {
    if (fs.existsSync(filePath)) {
        fs.unlinkSync(filePath);
    }
}

function normalizeLineEndings(text: string | undefined): string | undefined {
    return text ? text.replace(/\r\n|\r/g, "\n") : text;
}

function readFileIfExists(filePath: string): string | undefined {
    return normalizeLineEndings(fs.existsSync(filePath) ? fs.readFileSync(filePath).toString() : undefined);
}

describe("test cases", function() {
    const testCasesDirPath = path.resolve("./test/cases");
    const testCases = fs.readdirSync(testCasesDirPath);
    const repeatingFlag = process.env.REPEATING;

    testCases.forEach((testCaseDirName: string) => {
        ["ts", "js"].forEach((scriptType: string) => {
            const testCaseDirPath = path.resolve(testCasesDirPath, testCaseDirName);
            const sourceFileName = `functions.${scriptType}`;
            const sourceFile = path.resolve(testCaseDirPath, sourceFileName);
            const source: string | undefined = readFileIfExists(sourceFile);

            if (source) {
                it(`${testCaseDirName}\\${sourceFileName}`, async function() {
                    // add a file named "skip" to skip the test case
                    if (fs.existsSync(path.resolve(testCaseDirPath, "skip"))) {
                        this.skip();
                    // only run repeating parameter tests if the process variable REPEATING found
                    } else if (testCaseDirName.startsWith("repeating") && !repeatingFlag) {
                        this.skip();
                    } else {
                        const actualErrorsFile = path.join(testCaseDirPath, `actual.${scriptType}.errors.txt`);
                        const expectedErrorsFile = path.join(testCaseDirPath, `expected.${scriptType}.errors.txt`);
                        const actualMetadataFile = path.join(testCaseDirPath, `actual.${scriptType}.json`);
                        const expectedMetadataFile = path.join(testCaseDirPath, "expected.json");
                        const expectedMetadata: string | undefined = readFileIfExists(expectedMetadataFile);

                        // add a file named "debugger" to break on the test case
                        if (fs.existsSync(path.resolve(testCaseDirPath, "debugger"))) {
                            // tslint:disable-next-line: no-debugger
                            debugger;
                        }

                        // generate metadata
                        const result = await generate(sourceFile, actualMetadataFile);

                        const actualMetadata = readFileIfExists(actualMetadataFile);
                        const actualErrors = (result.errors.length > 0) ? result.errors.join("\n") : undefined;
                        const expectedErrors = readFileIfExists(expectedErrorsFile);

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
