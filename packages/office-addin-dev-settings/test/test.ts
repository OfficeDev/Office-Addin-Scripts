import * as assert from "assert";
import * as mocha from "mocha";
import * as devSettings from "../src/dev-settings";

const addinId = "9982ab78-55fb-472d-b969-b52ed294e173";

async function testSetSourceBundleUrlComponents(host?: string, port?: string, path?: string, extension?: string) {
  await devSettings.setSourceBundleUrl(addinId, new devSettings.SourceBundleUrlComponents(host, port, path, extension));
  const components: devSettings.SourceBundleUrlComponents = await devSettings.getSourceBundleUrl(addinId);
  assert.strictEqual(components.extension, extension);
  assert.strictEqual(components.host, host);
  assert.strictEqual(components.path, path);
  assert.strictEqual(components.port, port);
  assert.strictEqual(components.url, `http://${host || "localhost"}:${port || "8081"}/${path || "{path}"}${(extension === undefined) ? ".bundle" : extension}`);
}

describe("DevSettings", function() {
  this.beforeAll(function() {
    devSettings.clearDevSettings(addinId);
  });

  describe ("when no dev settings", function() {
    it("debugging should not be enabled", async function() {
      assert.strictEqual(await devSettings.isDebuggingEnabled(addinId), false);
    });
    it("live reload should not be enabled", async function() {
      assert.strictEqual(await devSettings.isLiveReloadEnabled(addinId), false);
    });
    it("have defaults for source bundle url components", async function() {
      const components: devSettings.SourceBundleUrlComponents = await devSettings.getSourceBundleUrl(addinId);
      assert.strictEqual(components.extension, undefined);
      assert.strictEqual(components.host, undefined);
      assert.strictEqual(components.path, undefined);
      assert.strictEqual(components.port, undefined);
      assert.strictEqual(components.url, "http://localhost:8081/{path}.bundle");
    });
    it("debugging can be enabled", async function() {
      assert.strictEqual(await devSettings.isDebuggingEnabled(addinId), false);
      await devSettings.enableDebugging(addinId);
      assert.strictEqual(await devSettings.isDebuggingEnabled(addinId), true);
    });
    it("live reload can be enabled", async function() {
      assert.strictEqual(await devSettings.isLiveReloadEnabled(addinId), false);
      await devSettings.enableLiveReload(addinId);
      assert.strictEqual(await devSettings.isLiveReloadEnabled(addinId), true);
    });
    it("source bundle url components can be set", async function() {
      await testSetSourceBundleUrlComponents("HOST", "PORT", "PATH", ".EXT");
    });
    it("source bundle url components can be cleared", async function() {
      await testSetSourceBundleUrlComponents(undefined, undefined, undefined, undefined);
    });
  });

  describe("when debugging is enabled", function() {
    it("debugging can be disabled", async function() {
      assert.strictEqual(await devSettings.isDebuggingEnabled(addinId), true);
      await devSettings.disableDebugging(addinId);
      assert.strictEqual(await devSettings.isDebuggingEnabled(addinId), false);
    });
  });

  describe("when debugging is not enabled", function() {
    it("debugging can be enabled", async function() {
      assert.strictEqual(await devSettings.isDebuggingEnabled(addinId), false);
      await devSettings.enableDebugging(addinId);
      assert.strictEqual(await devSettings.isDebuggingEnabled(addinId), true);
    });
  });

  describe("when live reload is enabled", function() {
    it("live reload can be disabled", async function() {
      assert.strictEqual(await devSettings.isLiveReloadEnabled(addinId), true);
      await devSettings.disableLiveReload(addinId);
      assert.strictEqual(await devSettings.isLiveReloadEnabled(addinId), false);
    });
  });

  describe("when live reload is not enabled", function() {
    it("live reload can be disabled", async function() {
      assert.strictEqual(await devSettings.isLiveReloadEnabled(addinId), false);
      await devSettings.enableLiveReload(addinId);
      assert.strictEqual(await devSettings.isLiveReloadEnabled(addinId), true);
    });
  });
});
