// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import assert from "assert";
import { describe, it } from "mocha";
import { getCacheTargets } from "../src/cache";

/* global process */

describe("getCacheTargets()", function () {
  it("returns an array", async function () {
    const targets = await getCacheTargets();
    assert.ok(Array.isArray(targets));
  });

  it("returns targets for the current platform", async function () {
    const targets = await getCacheTargets();
    if (process.platform === "win32" || process.platform === "darwin") {
      assert.ok(targets.length > 0, "expected at least one cache target on a supported platform");
      for (const target of targets) {
        assert.ok(typeof target.label === "string" && target.label.length > 0, "target label should be a non-empty string");
        assert.ok(typeof target.dir === "string" && target.dir.length > 0, "target dir should be a non-empty string");
      }
    } else {
      assert.strictEqual(targets.length, 0, "expected no targets on an unsupported platform");
    }
  });
});
