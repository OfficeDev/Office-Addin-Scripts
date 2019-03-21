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

function readFileIfExists(filePath: string): string | undefined {
    return (fs.existsSync(filePath)) ? fs.readFileSync(filePath).toString() : undefined;
}

describe("test cases", function() {
    const testCasesDirPath = path.resolve("./test/cases");
    const testCases = fs.readdirSync(testCasesDirPath, {withFileTypes: true});

    testCases.forEach((testCaseDir: fs.Dirent) => {
        const testCaseDirName = testCaseDir.name;
        ["ts", "js"].forEach((scriptType: string) => {
            const testCaseDirPath = path.resolve(testCasesDirPath, testCaseDirName);
            const sourceFileName = `functions.${scriptType}`;
            const sourceFile = path.resolve(testCaseDirPath, sourceFileName);
            const source: string | undefined = readFileIfExists(sourceFile);

            if (source) {
                it(`${testCaseDirName}\\${sourceFileName}`, async function() {
                    const actualErrorsFile = path.join(testCaseDirPath, `actual.${scriptType}.errors.txt`);
                    const expectedErrorsFile = path.join(testCaseDirPath, `expected.${scriptType}.errors.txt`);
                    const actualMetadataFile = path.join(testCaseDirPath, `actual.${scriptType}.json`);
                    const expectedMetadataFile = path.join(testCaseDirPath, "expected.json");
                    const expectedMetadata: string | undefined = readFileIfExists(expectedMetadataFile);

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

                    assert.strictEqual(actualMetadata, expectedMetadata, "verify metadata");
                    assert.strictEqual(actualErrors, expectedErrors, "verify errors");

                    // if actual metadata is what was expected, delete the actual metadata file
                    if (actualMetadata === expectedMetadata) {
                        deleteFileIfExists(actualMetadataFile);
                    }
                });
            }
        });
    });
});
