// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as assert from "assert";
import * as mocha from "mocha";
import * as sinon from "sinon";
import * as log from "../src/log";
import * as parse from "../src/parse";
import * as jsonScript from "../src/getPackageJsonScript";

function isError(err: Error, message: string): boolean {
  return (err instanceof Error) && err.message === message;
}

describe("office-addin-cli tests", function() {
  describe("log.ts", function() {
    describe("logErrorMessage()", function() {
      it("called with Error", function() {
        const spyConsoleError = sinon.spy(console, "error");
        const spyConsoleLog = sinon.spy(console, "log");

        const message = "This is an error.";
        const error = new Error(message);
        log.logErrorMessage(error);

        assert.ok(spyConsoleError.calledOnceWith(`Error: ${message}`));
        assert.ok(spyConsoleLog.notCalled);

        spyConsoleError.restore();
        spyConsoleLog.restore();
      });
      it("called with string", function() {
        const spyConsoleError = sinon.spy(console, "error");
        const spyConsoleLog = sinon.spy(console, "log");

        const message = "This is the error message.";
        log.logErrorMessage(message);

        assert.ok(spyConsoleError.calledOnceWith(`Error: ${message}`));
        assert.ok(spyConsoleLog.notCalled);

        spyConsoleError.restore();
        spyConsoleLog.restore();
      });
    });
  });

  describe("parse.ts", function() {
    const errorShouldBeANumber: Error = new Error("The value should be a number.");

    describe("parseNumber()", function() {
      it("number - integer", function() {
        assert.strictEqual(parse.parseNumber(100), 100);
      });
      it("number - decimal", function() {
        assert.strictEqual(parse.parseNumber(0.5), 0.5);
      });
      it("string - integer number", function() {
        assert.strictEqual(parse.parseNumber("100"), 100);
      });
      it("string - decimal number", function() {
        assert.strictEqual(parse.parseNumber("0.5"), 0.5);
      });
      it("string - empty", function() {
        assert.throws(() => parse.parseNumber(""), errorShouldBeANumber);
      });
      it("string - space", function() {
        assert.throws(() => parse.parseNumber(" "), errorShouldBeANumber);
      });
      it("undefined", function() {
        assert.strictEqual(parse.parseNumber(undefined), undefined);
      });
      it("null", function() {
        assert.throws(() => parse.parseNumber(null), errorShouldBeANumber);
      });
    });
  });

  describe("get_package_json_script.ts", function() {
    describe("getPackageJsonScript() for npmv6", function() {
      let npm_package_json_backup: string | undefined;
      before(async function() {
        npm_package_json_backup = process.env.npm_package_json;
      });
      it("npm - empty string", async function() {
        assert.strictEqual(await jsonScript.getPackageJsonScript(""), undefined);
      });
      it("npmv6 - exists", async function() {
        process.env["npm_package_scripts_test"] = "npmv6 1";
        process.env.npm_package_json = "";
        assert.strictEqual(await jsonScript.getPackageJsonScript("test"), "npmv6 1");
      });
      it("npmv6 - non existent", async function() {
        process.env.npm_package_json = "";
        assert.strictEqual(await jsonScript.getPackageJsonScript("npmv6testmissing"), undefined);
      });
      after(function() {
        process.env.npm_package_json = npm_package_json_backup;
      });
    });

    describe("getPackageJsonScript() for npmv7", function() {
      let npm_package_json_backup: string | undefined;
      before(async function() {
        npm_package_json_backup = process.env.npm_package_json;
      });
      it("npmv7 - no hyphen", async function() {
        process.env.npm_package_json = "./test/test.json";
        assert.strictEqual(await jsonScript.getPackageJsonScript("test"), "1");
      });
      it("npmv7 - hyphen", async function() {
        process.env.npm_package_json = "./test/test.json";
        assert.strictEqual(await jsonScript.getPackageJsonScript("test-test"), "2");
      });
      it("npmv7 - underscore", async function() {
        process.env.npm_package_json = "./test/test.json";
        assert.strictEqual(await jsonScript.getPackageJsonScript("test_test"), "3");
      });
      it("npmv7 - space", async function() {
        process.env.npm_package_json = "./test/test.json";
        assert.strictEqual(await jsonScript.getPackageJsonScript("test test"), "4");
      });
      it("npmv7 - multiple hyphens", async function() {
        process.env.npm_package_json = "./test/test.json";
        assert.strictEqual(await jsonScript.getPackageJsonScript("test-tes-te"), "5");
      });
      it("npmv7 - non existent", async function() {
        process.env.npm_package_json = "./test/test.json";
        assert.strictEqual(await jsonScript.getPackageJsonScript("testmissing"), undefined);
      });
      after(function() {
        process.env.npm_package_json = npm_package_json_backup;
      });
    });
  });
});
