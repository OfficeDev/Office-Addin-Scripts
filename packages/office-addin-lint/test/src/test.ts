// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as assert from "assert";
import * as fs from "fs";
import * as mocha from "mocha";
import * as path from "path";
import * as prettierLint from "../../src/lint";

describe("test cases", function() {
    it("eslint and prettier commands", async function() {
        const inputFile = "./test/cases/basic/functions.ts";
        const lintExpectedCommand = "eslint ./test/cases/basic/functions.ts";
        const lintFixExpectedCommand = "eslint --fix ./test/cases/basic/functions.ts";
        const prettierExpectedCommand = "prettier --parser typescript --write ./test/cases/basic/functions.ts";
        const lintCheckCommand = prettierLint.getLintCheckCommand(inputFile);
        assert.strictEqual(lintCheckCommand, lintExpectedCommand, "eslint lint command");
        const lintFixCommand = prettierLint.getLintFixCommand(inputFile);
        assert.strictEqual(lintFixCommand, lintFixExpectedCommand, "eslint fix command");
        const prettierCommand = prettierLint.getPrettierCommand(inputFile);
        assert.strictEqual(prettierCommand, prettierExpectedCommand, "prettier command");
    });
});
