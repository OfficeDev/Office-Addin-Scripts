// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import assert from "assert";
import fs from "fs";
import { describe, it } from "mocha";
import path from "path";
import { convertProject } from "../src/convert";
import { convert } from "office-addin-manifest-converter";

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
        const outputPath = path.join(process.env.TEMP as string, "ConvertManifstTest");
        await convert(manifestPath, outputPath);
        assert.strictEqual(fs.existsSync(path.join(outputPath, "manifest.json")), true);
      });
      it("Converts TaskPane manifest", async function() {
        this.timeout(6000);
        const manifestPath = "./test/TaskPane.manifest.xml";
        const outputPath = path.join(process.env.TEMP as string, "ConvertTaskpaneManifestTest");
        await convert(manifestPath, outputPath);
        assert.strictEqual(fs.existsSync(path.join(outputPath, "manifest.json")), true);
      });
      it("Can't convert malformed manifest", async function() {
        this.timeout(6000);
        const manifestPath = "./test/invalid.manifest.xml";
        const outputPath = path.join(process.env.TEMP as string, "ConvertMalformedManifestTest");
        convert(manifestPath, outputPath);
        assert.strictEqual(fs.existsSync(path.join(outputPath, "manifest.json")), false);
      });
    });
  });
});
