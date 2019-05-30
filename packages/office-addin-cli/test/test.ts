// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as assert from "assert";
import * as mocha from "mocha";
import * as sinon from "sinon";
import * as log from "../src/log";
import * as parse from "../src/parse";

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
});
