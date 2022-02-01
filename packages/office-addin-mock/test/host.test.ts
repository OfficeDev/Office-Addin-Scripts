// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as assert from "assert";
import * as mocha from "mocha";
import { getHostType, Host } from "../src/host";

describe("Test Host file", function() {
  describe("Works on differt hosts", function() {
    it("Excel", async function() {
      const object = {
        context: {},
        host: "excel"
      };
      assert.strictEqual(getHostType(object), "excel");
    });
    it("OneNote", async function() {
      const object = {
        context: {},
        host: "onenote"
      };
      assert.strictEqual(getHostType(object), "onenote");
    });
    it("PowerPoint", async function() {
      const object = {
        context: {},
        host: "powerpoint"
      };
      assert.strictEqual(getHostType(object), "powerpoint");
    });
    it("Project", async function() {
      const object = {
        context: {},
        host: "project"
      };
      assert.strictEqual(getHostType(object), "project");
    });
    it("Visio", async function() {
      const object = {
        context: {},
        host: "visio"
      };
      assert.strictEqual(getHostType(object), "visio");
    });
    it("Word", async function() {
      const object = {
        context: {},
        host: "word"
      };
      assert.strictEqual(getHostType(object), "word");
    });
    it("Outlook", async function() {
      const object = {
        context: {},
        host: "outlook"
      };
      assert.strictEqual(getHostType(object), "outlook");
    });
    it("Not defined", async function() {
      const outlookObject = {
        context: {},
      };
      assert.strictEqual(getHostType(outlookObject), "unknow");
    });
    it("Other", async function() {
      const object = {
        context: {},
        host: "not a host"

      };
      assert.strictEqual(getHostType(object), "unknow");
    });
  });
});
