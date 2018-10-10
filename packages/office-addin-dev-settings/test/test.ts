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

describe("DevSettingsForAddIn", function() {
  this.beforeAll(async function() {
    await devSettings.clearDevSettings(addinId);
  });

  this.afterAll (async function() {
    await devSettings.clearDevSettings(addinId);
  });

  describe("when no dev settings", function() {
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
    it("clear dev settings when no dev settings", async function() {
      await devSettings.clearDevSettings(addinId);
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

  describe("can specify debug method", function() {
    before("debugging should be disabled", async function() {
      await devSettings.disableDebugging(addinId);
      const methods = await devSettings.getEnabledDebuggingMethods(addinId);
      assert.strictEqual(methods.length, 0);
    }),
    it("direct debugging can be enabled", async function() {
      await devSettings.enableDebugging(addinId, true, devSettings.DebuggingMethod.Direct);
      const methods = await devSettings.getEnabledDebuggingMethods(addinId);
      assert.strictEqual(methods.length, 1);
      assert.strictEqual(methods[0], devSettings.DebuggingMethod.Direct);
    });
    it("web debugging can be enabled, and turns off direct debugging", async function() {
      await devSettings.enableDebugging(addinId, true, devSettings.DebuggingMethod.Web);
      const methods = await devSettings.getEnabledDebuggingMethods(addinId);
      assert.strictEqual(methods.length, 1);
      assert.strictEqual(methods[0], devSettings.DebuggingMethod.Web);
    });
    it("enabling direct debugging turns off web debugging", async function() {
      await devSettings.enableDebugging(addinId, true, devSettings.DebuggingMethod.Direct);
      const methods = await devSettings.getEnabledDebuggingMethods(addinId);
      assert.strictEqual(methods.length, 1);
      assert.strictEqual(methods[0], devSettings.DebuggingMethod.Direct);
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

describe("RuntimeLogging", function() {
  const defaultRuntimeLoggingPath = `${process.env.TEMP}\\OfficeAddins.log.txt`;
  const validRuntimeLoggingPath = `c:\\OfficeAddins.log.txt`;

  this.beforeAll(async function() {
    await devSettings.disableRuntimeLogging();
  });

  describe("when no runtime logging", function() {
    it("runtime logging should not be enabled", async function() {
      assert.strictEqual(await devSettings.getRuntimeLoggingPath(), undefined);
    });
    it("runtime logging can be enabled to the default path", async function() {
      assert.strictEqual(await devSettings.getRuntimeLoggingPath(), undefined);
      await devSettings.enableRuntimeLogging(undefined);
      assert.strictEqual(await devSettings.getRuntimeLoggingPath(), defaultRuntimeLoggingPath);
    });
    it("runtime logging can be disabled", async function() {
      await devSettings.disableRuntimeLogging();
      assert.strictEqual(await devSettings.getRuntimeLoggingPath(), undefined);
    });
    it("runtime logging can be enabled to the specified path", async function() {
      await devSettings.enableRuntimeLogging(validRuntimeLoggingPath);
      assert.strictEqual(await devSettings.getRuntimeLoggingPath(), validRuntimeLoggingPath);
    });
    it("runtime logging can be disabled", async function() {
      await devSettings.enableRuntimeLogging(validRuntimeLoggingPath);
      assert.strictEqual(await devSettings.getRuntimeLoggingPath(), validRuntimeLoggingPath);
    });
  });
});
