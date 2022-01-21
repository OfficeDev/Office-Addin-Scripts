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

const contextMockData = {
  workbook: {
    range: {
      address: "C2",
      format: {
        fill: {
          color: "green",
        }
      }
    },
    getSelectedRange: function () {
      return this.range;
    },
  },
};

describe("Test OfficeMockObject class", function() {
  describe("Populate object", function() {
    it("Object structure created", async function() {
      const officeMock = new OfficeMockObject(testObject);

      officeMock.range.load("color");
      await officeMock.sync();
      assert.strictEqual(officeMock.range.color, "blue");

      officeMock.range.font.load("size");
      await officeMock.sync();
      assert.strictEqual(officeMock.range.font.size, 12);

      assert.strictEqual(officeMock.notAProperty, undefined);
    });
    it("Context mock created and working", async function() {
      const contextMock = new OfficeMockObject(contextMockData);

      const range = contextMock.workbook.getSelectedRange();
      range.load("address");
      await contextMock.sync();

      assert.strictEqual(contextMock.workbook.range.address, "C2");
    });
    it("Multiple load calls", async function() {
      const officeMock = new OfficeMockObject(testObject);

      officeMock.range.load("color");
      await officeMock.sync();
      assert.strictEqual(officeMock.range.color, "blue");

      officeMock.range.load("color");
      await officeMock.sync();
      assert.strictEqual(officeMock.range.color, "blue");

      officeMock.range.load("color");
      assert.strictEqual(officeMock.range.color, "blue");

      await officeMock.sync();
      assert.strictEqual(officeMock.range.color, "blue");

      assert.strictEqual(officeMock.notAProperty, undefined);
    });
    it("Missing load", async function() {
      const officeMock = new OfficeMockObject(testObject);
      assert.strictEqual(officeMock.range.color, "Error, property was not loaded");
      await officeMock.sync();
      assert.strictEqual(officeMock.range.color, "Error, property was not loaded");
    });
    it("Missing sync", async function() {
      const officeMock = new OfficeMockObject(testObject);
      officeMock.range.load("color");
      assert.strictEqual(officeMock.range.color, "Error, context.sync() was not called");
      await officeMock.sync();
      officeMock.range.load("color");
      assert.strictEqual(officeMock.range.color, "blue");
    });
    it("Functions added", async function() {
      const officeMock = new OfficeMockObject(testObject);
      officeMock.range.load("color");
      await officeMock.sync();
      assert.strictEqual(officeMock.range.getColor(), "blue");
      officeMock.range.color = "green";
      assert.strictEqual(officeMock.range.getColor(), "green");
    });
    it("Writting values", async function() {
      const officeMock = new OfficeMockObject(testObject);

      officeMock.range.address = "C2";
      assert.strictEqual(officeMock.range.address, "C2");
  
      officeMock.range.font.color = "blue";
      assert.strictEqual(officeMock.range.font.color, "blue");
    });
  });

  describe("Writting values already present at object model", function() {
    it("Load and Sync", async function() {
      const officeMock = new OfficeMockObject(testObject);
      officeMock.range.color = "new color";
      officeMock.range.load("color");
      await officeMock.sync();
      assert.strictEqual(officeMock.range.color, "new color");
    });
    it("Only load", async function() {
      const officeMock = new OfficeMockObject(testObject);
      officeMock.range.color = "new color";
      officeMock.range.load("color");
      assert.strictEqual(officeMock.range.color, "new color");
    });
    it("Only sync", async function() {
      const officeMock = new OfficeMockObject(testObject);
      officeMock.range.color = "new color";
      await officeMock.sync();
      assert.strictEqual(officeMock.range.color, "new color");
    });
    it("No load and no sync", async function() {
      const officeMock = new OfficeMockObject(testObject);
      officeMock.range.color = "new color";
      assert.strictEqual(officeMock.range.color, "new color");
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
      await officeMock.sync();
      assert.strictEqual(officeMock.range.getColor(), "blue");
    });
    it("Navigational load", async function() {
      const officeMock = new OfficeMockObject(testObject);
      officeMock.load("range/color");
      await officeMock.sync();
      assert.strictEqual(officeMock.range.getColor(), "blue");

      officeMock.load("range/font/size");
      await officeMock.sync();
      assert.strictEqual(officeMock.range.font.size, 12);
  
      assert.throws(() => officeMock.load("range/notANavigational/size"));
      assert.throws(() => officeMock.load("notANavigational/font/size"));
    });
    it("Multiple properties load", async function() {
      const officeMock = new OfficeMockObject(testObject);
      officeMock.load("range/color, range/font/size");
      await officeMock.sync();
      assert.strictEqual(officeMock.range.getColor(), "blue");
      assert.strictEqual(officeMock.range.font.size, 12);
    });
    it("Comma separated load", async function() {
      const officeMock = new OfficeMockObject(testObject);
      officeMock.load(["range/color", "range/font/type", "range/font/size"]);
      await officeMock.sync();
      assert.strictEqual(officeMock.range.getColor(), "blue");
      assert.strictEqual(officeMock.range.font.type, "arial");
      assert.strictEqual(officeMock.range.font.size, 12);
    });
    it("Load star", async function() {
      const officeMock = new OfficeMockObject(testObject);
      officeMock.range.load("*");
      await officeMock.sync();
      assert.strictEqual(officeMock.range.getColor(), "blue");
    });
    it("Load navigational property", async function() {
      const contextMock = new OfficeMockObject(contextMockData);
      contextMock.workbook.range.load("format/fill");
      await contextMock.sync();
      assert.strictEqual(contextMock.workbook.range.format.fill.color, "green");
    });
    it("Load object property", async function() {
      const contextMock = new OfficeMockObject(contextMockData);
      contextMock.workbook.range.load({ format: { fill: { color: false } }, address: true } );
      await contextMock.sync();
      assert.strictEqual(contextMock.workbook.range.format.fill.color, "Error, property was not loaded");
      assert.strictEqual(contextMock.workbook.range.address, "C2");
      assert.throws(() => contextMock.load({ format: { notAProperty: false }, address: true } ));

      contextMock.workbook.range.load({ format: { fill: { color: true } } } );
      await contextMock.sync();
      assert.strictEqual(contextMock.workbook.range.format.fill.color, "green");
    });
    it("Loading an array", async function() {
      const ContextMockData = {
        items: [ "text", "text2" ],
      };
      const context = new OfficeMockObject(ContextMockData) as any;
      context.load("items");
      await context.sync();

      assert.deepStrictEqual(context.items, [ "text", "text2" ]);
    });
    it("Loading an array of objects", async function() {
      const ContextMockData = {
        items: [ { text: 'A' }, {text2: 'B' } ],
      };
  
      const context = new OfficeMockObject(ContextMockData) as any;
      context.load( { items: [ { text: 'A' }, { text2: 'B' } ] } );
      await context.sync();

      assert.strictEqual(context.items[0].text, "A");
      assert.strictEqual(context.items[1].text2, "B");
    });
  });

  describe("Works on Outlook", function() {
    it("Object construction", async function() {
      const officeMock = new OfficeMockObject(testObject, true);
      officeMock.load("range");
      officeMock.sync();
      assert.strictEqual(officeMock.range.color, "blue");
      assert.strictEqual(officeMock.range.font.size, 12);
      assert.strictEqual(officeMock.range.getColor(), "blue");
    });
    it("Invalid load calls", async function() {
      const officeMock = new OfficeMockObject(testObject, true);
      officeMock.range.load("color");
      assert.strictEqual(officeMock.range.getColor(), "blue");
      assert.strictEqual(officeMock.range.color, "blue");
    });
    it("Invalid sync calls", async function() {
      const officeMock = new OfficeMockObject(testObject, true);
      officeMock.sync();
      assert.strictEqual(officeMock.range.getColor(), "blue");
      assert.strictEqual(officeMock.range.color, "blue");
    });
  });
});
