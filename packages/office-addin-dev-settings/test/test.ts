import * as assert from "assert";
import * as fsextra from "fs-extra";
import * as mocha from "mocha";
import * as fspath from "path";
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

describe("RuntimeLogging", async function() {
  const testExecDirName = "testExec";
  const defaultFileName = "OfficeAddins.log.txt";
  const baseDirPath = process.cwd();
  const testExecDirPath = fspath.join(baseDirPath, testExecDirName);
  const defaultPath = `${process.env.TEMP}\\${defaultFileName}`;

  this.beforeAll(async function() {
    await devSettings.disableRuntimeLogging();
  });

  describe("basic validation", function() {
    it("runtime logging should not be enabled", async function() {
      assert.strictEqual(await devSettings.getRuntimeLoggingPath(), undefined);
    });
    it("runtime logging can be enabled to the default path", async function() {
      assert.strictEqual(await devSettings.getRuntimeLoggingPath(), undefined);
      await devSettings.enableRuntimeLogging(undefined);
      assert.strictEqual(await devSettings.getRuntimeLoggingPath(), defaultPath);
    });
    it("runtime logging can be disabled", async function() {
      await devSettings.disableRuntimeLogging();
      assert.strictEqual(await devSettings.getRuntimeLoggingPath(), undefined);
    });
  });

  describe("enableRuntimeLogging", function() {
    it("default path (TEMP folder exists)", async function() {
      assert.strictEqual(await devSettings.getRuntimeLoggingPath(), undefined);
      const path: string = await devSettings.enableRuntimeLogging(undefined);
      assert.strictEqual(path, defaultPath);
      assert.strictEqual(path, await devSettings.getRuntimeLoggingPath());
    });
    it("default path but TEMP folder doesn't exist", async function() {
      const TEMP = process.env.TEMP;
      const directoryDoesNotExist = fspath.join(testExecDirName, "doesNotExist", defaultFileName);
      process.env.TEMP = directoryDoesNotExist;
      let error;
      try {
        const path: string = await devSettings.enableRuntimeLogging();
      } catch (err) {
        error = err;
      }
      assert.ok(error instanceof Error, "should throw an error");
      assert.strictEqual(error.message, `You need to specify the path where the file can be written. Unable to write to: "${directoryDoesNotExist}\\${defaultFileName}".`);
      process.env.TEMP = TEMP;
    });
    it("default path but TEMP env var is not defined", async function() {
      const env = process.env;
      process.env = { };
      let error;
      try {
        const path: string = await devSettings.enableRuntimeLogging();
      } catch (err) {
        error = err;
      }
      assert.ok(error instanceof Error, "should throw an error");
      assert.strictEqual(error.message, "The TEMP environment variable is not defined.");
      process.env = env;
    });

    describe("filesystem test cases", async function() {
      this.beforeEach(async function() {
        await fsextra.remove(testExecDirPath);
        await fsextra.mkdir(testExecDirPath);
      });
      this.afterAll(async function() {
        await fsextra.remove(testExecDirPath);
      });
      it("directory does not exist", async function() {
        const filePath = fspath.join(testExecDirPath, "doesNotExist", defaultFileName);
        let error;
        try {
          const path = await devSettings.enableRuntimeLogging(filePath);
        } catch (err) {
          error = err;
        }
        assert.ok(error instanceof Error, "should throw an error");
        assert.strictEqual(error.message, `You need to specify the path where the file can be written. Unable to write to: "${filePath}".`);
        });
      it("file does not exist in writable directory", async function() {
        try {
          const filePath = fspath.join(testExecDirPath, defaultFileName);
          const path = await devSettings.enableRuntimeLogging(filePath);
          assert.strictEqual(path, filePath);
          assert.strictEqual(path, await devSettings.getRuntimeLoggingPath());
        } catch (err) {
          assert.fail("should not throw an error");
        }
      });
      // it("file does not exist in non-writable directory", async function() {
      //   const subdirPath = fspath.join(testExecDirPath, "subdir");
      //   const filePath = fspath.join(subdirPath, defaultFileName);
      //   let error;

      //   // create the subdir
      //   fsextra.mkdir(subdirPath);

      //   // make the subdir non-writable
      //   fsextra.chmod(subdirPath, 0o444);

      //   try {
      //     const path = await devSettings.enableRuntimeLogging(filePath);
      //   } catch (err) {
      //     error = err;
      //   }

      //   // clean up
      //   fsextra.chmod(subdirPath, 0o777);

      //   assert.ok(error instanceof Error, "should throw an error");
      //   assert.strictEqual(error.message, `You need to specify the path where the file can be written. Unable to write to: "${filePath}".`);
      // });
      it("file already exists and is writable", async function() {
        try {
          const filePath = fspath.join(testExecDirPath, defaultFileName);

          // create the file
          const file = await fsextra.open(filePath, "a+");
          await fsextra.close(file);

          const path = await devSettings.enableRuntimeLogging(filePath);
          assert.strictEqual(path, filePath);
          assert.strictEqual(path, await devSettings.getRuntimeLoggingPath());
        } catch (err) {
          assert.fail("should not throw an error");
        }
      });
      it("file already exists but is not writable", async function() {
        const filePath = fspath.join(testExecDirPath, defaultFileName);
        let error;
        try {
          // create the file
          const file = await fsextra.open(filePath, "a+");
          await fsextra.close(file);

          // make the file read-only
          await fsextra.chmod(filePath, 0o444);

          const path = await devSettings.enableRuntimeLogging(filePath);
          assert.strictEqual(path, filePath);
          assert.strictEqual(path, await devSettings.getRuntimeLoggingPath());
        } catch (err) {
          error = err;
        }
        assert.ok(error instanceof Error, "should throw an error");
        assert.strictEqual(error.message, `You need to specify the path to a writable file. Unable to write to: "${filePath}".`);
      });
      // it("file already exists but directory is not writable", async function() {
      //   const subdirPath = fspath.join(testExecDirPath, "subdir");
      //   const filePath = fspath.join(subdirPath, defaultFileName);
      //   let error;

      //   // create the subdir
      //   fsextra.mkdir(subdirPath);

      //   // create the file
      //   const file = await fsextra.open(filePath, "a+");
      //   await fsextra.close(file);

      //   // make the subdir non-writable
      //   fsextra.chmod(subdirPath, 0o444);

      //   try {
      //     const path = await devSettings.enableRuntimeLogging(filePath);
      //   } catch (err) {
      //     error = err;
      //   }

      //   // clean up
      //   fsextra.chmod(subdirPath, 0o777);

      //   assert.ok(error instanceof Error, "should throw an error");
      //   assert.strictEqual(error.message, `You need to specify the path where the file can be written. Unable to write to: "${filePath}".`);
      // });
    });
  });
});
