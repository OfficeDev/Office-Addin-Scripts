import * as assert from "assert";
import * as mocha from "mocha";
import * as devSettings from "../src/dev-settings";

const addinId = "9982ab78-55fb-472d-b969-b52ed294e173";

describe("DevSettings", function() {
  this.beforeAll(function() {
    devSettings.clearDevSettings(addinId);
  });

  describe ("when no dev settings", function() {
    it("debugging should not be enabled", async function() {
      assert.equal(await devSettings.isDebuggingEnabled(addinId), false);
    });
    it("live reload should not be enabled", async function() {
      assert.equal(await devSettings.isLiveReloadEnabled(addinId), false);
    });
    it("debugging can be enabled", async function() {
      assert.equal(await devSettings.isDebuggingEnabled(addinId), false);
      await devSettings.enableDebugging(addinId);
      assert.equal(await devSettings.isDebuggingEnabled(addinId), true);
    });
    it("live reload can be enabled", async function() {
      assert.equal(await devSettings.isLiveReloadEnabled(addinId), false);
      await devSettings.enableLiveReload(addinId);
      assert.equal(await devSettings.isLiveReloadEnabled(addinId), true);
    });
  });

  describe("when debugging is enabled", function() {
    it("debugging can be disabled", async function() {
      assert.equal(await devSettings.isDebuggingEnabled(addinId), true);
      await devSettings.disableDebugging(addinId);
      assert.equal(await devSettings.isDebuggingEnabled(addinId), false);
    });
  });

  describe("when debugging is not enabled", function() {
    it("debugging can be enabled", async function() {
      assert.equal(await devSettings.isDebuggingEnabled(addinId), false);
      await devSettings.enableDebugging(addinId);
      assert.equal(await devSettings.isDebuggingEnabled(addinId), true);
    });
  });

  describe("when live reload is enabled", function() {
    it("live reload can be disabled", async function() {
      assert.equal(await devSettings.isLiveReloadEnabled(addinId), true);
      await devSettings.disableLiveReload(addinId);
      assert.equal(await devSettings.isLiveReloadEnabled(addinId), false);
    });
  });

  describe("when live reload is not enabled", function() {
    it("live reload can be disabled", async function() {
      assert.equal(await devSettings.isLiveReloadEnabled(addinId), false);
      await devSettings.enableLiveReload(addinId);
      assert.equal(await devSettings.isLiveReloadEnabled(addinId), true);
    });
  });
});
