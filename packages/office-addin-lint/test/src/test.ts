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
        const lintExpectedCommand = "./test/cases/basic/functions.ts";
        const lintTestExpectedCommand = ".eslintrc.test.json"
        const lintFixExpectedCommand = "--fix ./test/cases/basic/functions.ts";
        const prettierExpectedCommand = "--parser typescript --write ./test/cases/basic/functions.ts";

        const lintCheckCommand = prettierLint.getLintCheckCommand(inputFile);
        assert.strictEqual(lintCheckCommand.indexOf(lintExpectedCommand) > 0 , true, "Lint command does not match expected value.");

        const lintCheckTestCommand = prettierLint.getLintCheckCommand(inputFile, true);
        assert.strictEqual(lintCheckTestCommand.indexOf(lintTestExpectedCommand) > 0 , true, "Lint test command does not match expected value.");

        const lintFixCommand = prettierLint.getLintFixCommand(inputFile);
        assert.strictEqual(lintFixCommand.indexOf(lintFixExpectedCommand) > 0, true, "Lint fix command does not match expected value.");

        const prettierCommand = prettierLint.getPrettierCommand(inputFile);
        assert.strictEqual(prettierCommand.indexOf(prettierExpectedCommand) > 0, true, "Prettier command does not match expected value.");
    });

    it("spaces in filename", async function() {
        const inputFile = "./test/cases/basic/functions with space.ts";
        const lintExpectedCommand = "./test/cases/basic/functions\\ with\\ space.ts";
        const lintFixExpectedCommand = "--fix ./test/cases/basic/functions\\ with\\ space.ts";
        const prettierExpectedCommand = "--parser typescript --write ./test/cases/basic/functions\\ with\\ space.ts";

        const lintCheckCommand = prettierLint.getLintCheckCommand(inputFile);
        assert.strictEqual(lintCheckCommand.indexOf(lintExpectedCommand) > 0 , true, "Lint command does not match expected value.");

        const lintFixCommand = prettierLint.getLintFixCommand(inputFile);
        assert.strictEqual(lintFixCommand.indexOf(lintFixExpectedCommand) > 0, true, "Lint fix command does not match expected value.");

        const prettierCommand = prettierLint.getPrettierCommand(inputFile);
        assert.strictEqual(prettierCommand.indexOf(prettierExpectedCommand) > 0, true, "Prettier command does not match expected value.");
    });

    it("spaces in filepath", async function() {
        const inputFile = "./test/cases/basic with space/functions.ts";
        const lintExpectedCommand = "./test/cases/basic\\ with\\ space/functions.ts";
        const lintFixExpectedCommand = "--fix ./test/cases/basic\\ with\\ space/functions.ts";
        const prettierExpectedCommand = "--parser typescript --write ./test/cases/basic\\ with\\ space/functions.ts";

        const lintCheckCommand = prettierLint.getLintCheckCommand(inputFile);
        assert.strictEqual(lintCheckCommand.indexOf(lintExpectedCommand) > 0 , true, "Lint command does not match expected value.");

        const lintFixCommand = prettierLint.getLintFixCommand(inputFile);
        assert.strictEqual(lintFixCommand.indexOf(lintFixExpectedCommand) > 0, true, "Lint fix command does not match expected value.");

        const prettierCommand = prettierLint.getPrettierCommand(inputFile);
        assert.strictEqual(prettierCommand.indexOf(prettierExpectedCommand) > 0, true, "Prettier command does not match expected value.");
    });
});
