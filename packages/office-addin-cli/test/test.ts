// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

/// <reference types="../src/read-package-json-fast"/>

import assert from "assert";
import { afterEach, beforeEach, describe, it } from "mocha";
import * as parse from "../src/parse";
import { clearCachedScripts, getPackageJsonScript } from "../src/npmPackage";

/* global process */

describe("office-addin-cli tests", function () {
  describe("parse.ts", function () {
    const errorShouldBeANumber: Error = new Error("The value should be a number.");

    describe("parseNumber()", function () {
      it("number - integer", function () {
        assert.strictEqual(parse.parseNumber(100), 100);
      });
      it("number - decimal", function () {
        assert.strictEqual(parse.parseNumber(0.5), 0.5);
      });
      it("string - integer number", function () {
        assert.strictEqual(parse.parseNumber("100"), 100);
      });
      it("string - decimal number", function () {
        assert.strictEqual(parse.parseNumber("0.5"), 0.5);
      });
      it("string - empty", function () {
        assert.throws(() => parse.parseNumber(""), errorShouldBeANumber);
      });
      it("string - space", function () {
        assert.throws(() => parse.parseNumber(" "), errorShouldBeANumber);
      });
      it("undefined", function () {
        assert.strictEqual(parse.parseNumber(undefined), undefined);
      });
      it("null", function () {
        assert.throws(() => parse.parseNumber(null), errorShouldBeANumber);
      });
    });
  });

  describe("getPackageJsonScript.ts", function () {
    describe("getPackageJsonScript()", function () {
      const envBackup: { [key: string]: string | undefined } = {};
      beforeEach(async function () {
        envBackup.npm_package_json = process.env.npm_package_json;
      });
      afterEach(async function () {
        clearCachedScripts();

        if (envBackup.npm_package_json) {
          process.env.npm_package_json = envBackup.npm_package_json;
        } else {
          delete process.env.npm_package_json;
        }
        delete process.env.npm_package_scripts_one;
        delete process.env.npm_package_scripts_two_two;
        delete process.env.npm_package_scripts_three_three;
        delete process.env.npm_package_scripts_four_four;
        delete process.env.npm_package_scripts_five_five_five;
      });
      it("NPM v6: script name empty", async function () {
        delete process.env.npm_package_json;
        const script = await getPackageJsonScript("");
        assert.strictEqual(script, undefined);
      });
      it("NPM v6: script name exists", async function () {
        delete process.env.npm_package_json;
        process.env.npm_package_scripts_one = "npmv6 1";
        const script = await getPackageJsonScript("one");
        assert.strictEqual(script, "npmv6 1");
      });
      it("NPM v6: script name does not exist", async function () {
        delete process.env.npm_package_json;
        const script = await getPackageJsonScript("npmv6testmissing");
        assert.strictEqual(script, undefined);
      });
      it("NPM v6: script name contains hyphen", async function () {
        delete process.env.npm_package_json;
        process.env.npm_package_scripts_two_two = "2";
        const script = await getPackageJsonScript("two-two");
        assert.strictEqual(script, "2");
      });
      it("NPM v6: script name contains underscore", async function () {
        delete process.env.npm_package_json;
        process.env.npm_package_scripts_three_three = "3";
        const script = await getPackageJsonScript("three_three");
        assert.strictEqual(script, "3");
      });
      it("NPM v6: script name contains multiple hyphens", async function () {
        delete process.env.npm_package_json;
        process.env.npm_package_scripts_five_five_five = "5";
        const script = await getPackageJsonScript("five-five-five");
        assert.strictEqual(script, "5");
      });
      it("NPM v7: script exists", async function () {
        process.env.npm_package_json = "./test/test.json";
        const script = await getPackageJsonScript("one");
        assert.strictEqual(script, "1");
      });
      it("NPM v7: script name does not exist", async function () {
        process.env.npm_package_json = "./test/test.json";
        const script = await getPackageJsonScript("testmissing");
        assert.strictEqual(script, undefined);
      });
      it("NPM v7: script name contains hyphen", async function () {
        process.env.npm_package_json = "./test/test.json";
        const script = await getPackageJsonScript("two-two");
        assert.strictEqual(script, "2");
      });
      it("NPM v7: script name contains underscore", async function () {
        process.env.npm_package_json = "./test/test.json";
        const script = await getPackageJsonScript("three_three");
        assert.strictEqual(script, "3");
      });
      it("NPM v7: script name contains a space", async function () {
        process.env.npm_package_json = "./test/test.json";
        const script = await getPackageJsonScript("four four");
        assert.strictEqual(script, "4");
      });
      it("NPM v7: script name contains multiple hyphens", async function () {
        process.env.npm_package_json = "./test/test.json";
        const script = await getPackageJsonScript("five-five-five");
        assert.strictEqual(script, "5");
      });
    });
  });
});
