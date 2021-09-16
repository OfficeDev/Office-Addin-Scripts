// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as assert from "assert";
import * as mocha from "mocha";
import { OfficeJSMock } from "../src/main";

const testObject = {
  "range": {
    "color": "blue",
    "getColor": function() {
      return this.color;
    },
    "font": {
      "size": 12
    }
  },
}

describe("OfficeJSMock", function() {
  describe("Populate with Object", function() {
    it("Object structure created", async function() {
      const officeJSMock = new OfficeJSMock(testObject);

      officeJSMock.range.load("color");
      officeJSMock.sync();
      assert.strictEqual(officeJSMock.range.color, "blue");

      officeJSMock.range.font.load("size");
      officeJSMock.sync();
      assert.strictEqual(officeJSMock.range.font.size, 12);

      assert.strictEqual(officeJSMock.notAProperty, undefined);
    });
    it("Missing load", async function() {
      const officeJSMock = new OfficeJSMock(testObject);
      assert.strictEqual(officeJSMock.range.color, "Error, property was not loaded");
      officeJSMock.sync();
      assert.strictEqual(officeJSMock.range.color, "Error, property was not loaded");
    });
    it("Missing sync", async function() {
      const officeJSMock = new OfficeJSMock(testObject);
      officeJSMock.range.load("color");
      assert.strictEqual(officeJSMock.range.color, "Error, context.sync() was not called");
      officeJSMock.sync();
      officeJSMock.range.load("color");
      assert.strictEqual(officeJSMock.range.color, "Error, context.sync() was not called");
    });
    it("Functions added", async function() {
      const officeJSMock = new OfficeJSMock(testObject);
      officeJSMock.range.load("color");
      officeJSMock.sync();
      assert.strictEqual(officeJSMock.range.getColor(), "blue");
      officeJSMock.range.color = "green";
      assert.strictEqual(officeJSMock.range.getColor(), "green");

      officeJSMock.range.setMock("color", "yellow");
      officeJSMock.range.load("color");
      officeJSMock.sync();
      assert.strictEqual(officeJSMock.range.getColor(), "yellow");
    });
  });

  describe("Load", function() {
    it("Invalid load", async function() {
      const officeJSMock = new OfficeJSMock(testObject);
      assert.throws(() => officeJSMock.range.load("notAProperty"));
    });
    it("Load on navigation property", async function() {
      const officeJSMock = new OfficeJSMock(testObject);
      officeJSMock.load("range");
      officeJSMock.sync();
      assert.strictEqual(officeJSMock.range.getColor(), "blue");
    });
    it("Navigational load", async function() {
      const officeJSMock = new OfficeJSMock(testObject);
      officeJSMock.load("range/color");
      officeJSMock.sync();
      assert.strictEqual(officeJSMock.range.getColor(), "blue");

      officeJSMock.load("range/font/size");
      officeJSMock.sync();
      assert.strictEqual(officeJSMock.range.font.size, 12);
  
      assert.throws(() => officeJSMock.load("range/notANavigational/size"));
      assert.throws(() => officeJSMock.load("notANavigational/font/size"));
    });
    it("Multiple properties load", async function() {
      const officeJSMock = new OfficeJSMock(testObject);
      officeJSMock.load("range/color, range/font/size");
      officeJSMock.sync();
      assert.strictEqual(officeJSMock.range.getColor(), "blue");
      assert.strictEqual(officeJSMock.range.font.size, 12);
    });
  });
});
