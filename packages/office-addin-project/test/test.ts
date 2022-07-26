// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as assert from "assert";
import * as fs from "fs";
import * as mocha from "mocha";
import * as path from "path";
import { convertManifest, convertProject } from "../src/convert";
import { ExpectedError } from "office-addin-usage-data";

describe("office-addin-project tests", function() {
  describe("convert.ts", function() {
    describe("convertProject", function() {
      it("Throws when manifest file does not exist", async function() {
        try {
          await convertProject("foo/bar.xml");
          assert.fail("The expected Error was not thrown.");
        } catch (err: any) {}
      });
      it("Throws when coverting already converted project", async function() {
        try {
          await convertProject("test/test.json");
          assert.fail("The expected Error was not thrown.");
        } catch (err: any) {}
      });
    });
    describe("convertManifest", function() {
      it("Converts test manifest", async function() {
        this.timeout(6000);
        const manifestPath = "./test/test-manifest.xml";
        const outputPath = "./temp/";
        convertManifest(manifestPath, outputPath);
        assert.strictEqual(fs.existsSync(path.join(outputPath, "test-manifest.json")), true);
      });
      it("Can't convert malformed manifest", async function() {
        this.timeout(6000);
        const manifestPath = "./test/invalid.manifest.xml";
        const outputPath = "./out";
        convertManifest(manifestPath, outputPath);
        assert.strictEqual(fs.existsSync(path.join(outputPath, "manifest.json")), false);
      });
    });
  });
});
