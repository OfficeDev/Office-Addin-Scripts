// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as assert from "assert";
import * as mocha from "mocha";
import { convertProject } from "../src/convert";
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
  });
});
