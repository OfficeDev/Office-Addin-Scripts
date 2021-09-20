// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as assert from "assert";
import * as mocha from "mocha";
import { OfficeMockObject } from "../src/main";

const testObject = {
  range: {
    color: "blue",
    getColor: function() {
      return this.color;
    },
    font: {
      size: 12,
      type: "arial"
    }
  },
}

describe("Test OfficeMockObject class", function() {
  describe("Populate object", function() {
    it("Object structure created", async function() {
      const officeMock = new OfficeMockObject(testObject);

      officeMock.range.load("color");
      officeMock.sync();
      assert.strictEqual(officeMock.range.color, "blue");

      officeMock.range.font.load("size");
      officeMock.sync();
      assert.strictEqual(officeMock.range.font.size, 12);

      assert.strictEqual(officeMock.notAProperty, undefined);
    });
    it("Multiple load calls", async function() {
      const officeMock = new OfficeMockObject(testObject);

      officeMock.range.load("color");
      officeMock.sync();
      assert.strictEqual(officeMock.range.color, "blue");

      officeMock.range.load("color");
      officeMock.sync();
      assert.strictEqual(officeMock.range.color, "blue");

      officeMock.range.load("color");
      assert.strictEqual(officeMock.range.color, "blue");

      officeMock.sync();
      assert.strictEqual(officeMock.range.color, "blue");

      assert.strictEqual(officeMock.notAProperty, undefined);
    });
    it("Missing load", async function() {
      const officeMock = new OfficeMockObject(testObject);
      assert.strictEqual(officeMock.range.color, "Error, property was not loaded");
      officeMock.sync();
      assert.strictEqual(officeMock.range.color, "Error, property was not loaded");
    });
    it("Missing sync", async function() {
      const officeMock = new OfficeMockObject(testObject);
      officeMock.range.load("color");
      assert.strictEqual(officeMock.range.color, "Error, context.sync() was not called");
      officeMock.sync();
      officeMock.range.load("color");
      assert.strictEqual(officeMock.range.color, "blue");
    });
    it("Functions added", async function() {
      const officeMock = new OfficeMockObject(testObject);
      officeMock.range.load("color");
      officeMock.sync();
      assert.strictEqual(officeMock.range.getColor(), "blue");
      officeMock.range.color = "green";
      assert.strictEqual(officeMock.range.getColor(), "green");

      officeMock.range.setMock("color", "yellow");
      officeMock.range.load("color");
      officeMock.sync();
      assert.strictEqual(officeMock.range.getColor(), "yellow");
    });
  });

  describe("Different ways to load properties", function() {
    it("Invalid load", async function() {
      const officeMock = new OfficeMockObject(testObject);
      assert.throws(() => officeMock.range.load("notAProperty"));
    });
    it("Load on navigation property", async function() {
      const officeMock = new OfficeMockObject(testObject);
      officeMock.load("range");
      officeMock.sync();
      assert.strictEqual(officeMock.range.getColor(), "blue");
    });
    it("Navigational load", async function() {
      const officeMock = new OfficeMockObject(testObject);
      officeMock.load("range/color");
      officeMock.sync();
      assert.strictEqual(officeMock.range.getColor(), "blue");

      officeMock.load("range/font/size");
      officeMock.sync();
      assert.strictEqual(officeMock.range.font.size, 12);
  
      assert.throws(() => officeMock.load("range/notANavigational/size"));
      assert.throws(() => officeMock.load("notANavigational/font/size"));
    });
    it("Multiple properties load", async function() {
      const officeMock = new OfficeMockObject(testObject);
      officeMock.load("range/color, range/font/size");
      officeMock.sync();
      assert.strictEqual(officeMock.range.getColor(), "blue");
      assert.strictEqual(officeMock.range.font.size, 12);
    });
    it("Comma separated load", async function() {
      const officeMock = new OfficeMockObject(testObject);
      officeMock.load(["range/color", "range/font/type", "range/font/size"]);
      officeMock.sync();
      assert.strictEqual(officeMock.range.getColor(), "blue");
      assert.strictEqual(officeMock.range.font.type, "arial");
      assert.strictEqual(officeMock.range.font.size, 12);
    });
  });
});
