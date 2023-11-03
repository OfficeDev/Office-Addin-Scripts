// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as assert from "assert";
import * as fsextra from "fs-extra";
import { OfficeApp, ManifestInfo, OfficeAddinManifest, AddInType } from "office-addin-manifest";
import * as os from "os";
import * as fspath from "path";
import * as sinon from "sinon";
import * as appcontainer from "../../src/appcontainer";
import { AppType } from "../../src/appType";
import * as devSettings from "../../src/dev-settings";
import * as devSettingsWindows from "../../src/dev-settings-windows";
import { registerWithTeams } from "../../src/publish";
import { deleteKey, getStringValue } from "../../src/registry";
import * as devSettingsSideload from "../../src/sideload";
const addinId = "9982ab78-55fb-472d-b969-b52ed294e173";
const isWindows = (process.platform === "win32");
const isMac = (process.platform === "darwin");

async function testSetSourceBundleUrlComponents(components: devSettings.SourceBundleUrlComponents, expected: devSettings.SourceBundleUrlComponents) {
  await devSettings.setSourceBundleUrl(addinId, components);
  const actual: devSettings.SourceBundleUrlComponents = await devSettings.getSourceBundleUrl(addinId);
  assert.strictEqual(actual.extension, expected.extension);
  assert.strictEqual(actual.host, expected.host);
  assert.strictEqual(actual.path, expected.path);
  assert.strictEqual(actual.port, expected.port);
  assert.strictEqual(actual.url, `http://${expected.host || "localhost"}:${expected.port || "8081"}/${expected.path || "{path}"}${expected.extension || ".bundle"}`);
}

describe("DevSettingsForAddIn", function() {
  if (isWindows) {
    this.beforeAll(async function() {
      await devSettings.clearDevSettings(addinId);
    });

    this.afterAll(async function() {
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
      it("webView override should not be set", async function() {
        assert.strictEqual(await devSettings.getWebView(addinId), undefined);
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
        const actual = new devSettings.SourceBundleUrlComponents("HOST", "PORT", "PATH", ".EXT");
        const expected = new devSettings.SourceBundleUrlComponents("HOST", "PORT", "PATH", ".EXT");
        await testSetSourceBundleUrlComponents(actual, expected);
      });
      it("source bundle url components can be cleared", async function() {
        const actual = new devSettings.SourceBundleUrlComponents("", "", "", "");
        const expected = new devSettings.SourceBundleUrlComponents(undefined, undefined, undefined, undefined);
        await testSetSourceBundleUrlComponents(actual, expected);
      });
      it("source bundle url host only can be set", async function() {
        const actual = new devSettings.SourceBundleUrlComponents("HOST", undefined, undefined, undefined);
        const expected = new devSettings.SourceBundleUrlComponents("HOST", undefined, undefined, undefined);
        await testSetSourceBundleUrlComponents(actual, expected);
      });
      it("source bundle url port only can be set", async function() {
        const actual = new devSettings.SourceBundleUrlComponents(undefined, "9999", undefined, undefined);
        const expected = new devSettings.SourceBundleUrlComponents("HOST", "9999", undefined, undefined);
        await testSetSourceBundleUrlComponents(actual, expected);
      });
      it("source bundle url path only can be set", async function() {
        const actual = new devSettings.SourceBundleUrlComponents(undefined, undefined, "PATH", undefined);
        const expected = new devSettings.SourceBundleUrlComponents("HOST", "9999", "PATH", undefined);
        await testSetSourceBundleUrlComponents(actual, expected);
      });
      it("source bundle url path only can be set", async function() {
        const actual = new devSettings.SourceBundleUrlComponents(undefined, undefined, undefined, "EXT");
        const expected = new devSettings.SourceBundleUrlComponents("HOST", "9999", "PATH", "EXT");
        await testSetSourceBundleUrlComponents(actual, expected);
      });
      it("source bundle url host only can be cleared", async function() {
        const actual = new devSettings.SourceBundleUrlComponents("", undefined, undefined, undefined);
        const expected = new devSettings.SourceBundleUrlComponents(undefined, "9999", "PATH", "EXT");
        await testSetSourceBundleUrlComponents(actual, expected);
      });
      it("source bundle url port only can be set", async function() {
        const actual = new devSettings.SourceBundleUrlComponents(undefined, "", undefined, undefined);
        const expected = new devSettings.SourceBundleUrlComponents(undefined, undefined, "PATH", "EXT");
        await testSetSourceBundleUrlComponents(actual, expected);
      });
      it("source bundle url path only can be set", async function() {
        const actual = new devSettings.SourceBundleUrlComponents(undefined, undefined, "", undefined);
        const expected = new devSettings.SourceBundleUrlComponents(undefined, undefined, undefined, "EXT");
        await testSetSourceBundleUrlComponents(actual, expected);
      });
      it("source bundle url path only can be set", async function() {
        const actual = new devSettings.SourceBundleUrlComponents(undefined, undefined, undefined, "");
        const expected = new devSettings.SourceBundleUrlComponents(undefined, undefined, undefined, undefined);
        await testSetSourceBundleUrlComponents(actual, expected);
      });
      it("webView can be set to IE", async function() {
        await devSettings.setWebView(addinId, devSettings.WebViewType.IE);
        assert.strictEqual(await devSettingsWindows.getWebView(addinId), "IE");
      });
      it("webView can be set to Edge", async function() {
        await devSettings.setWebView(addinId, devSettings.WebViewType.Edge);
        assert.strictEqual(await devSettingsWindows.getWebView(addinId), "Edge");
      });
      it("webView can be set to Edge Chromium", async function() {
        await devSettings.setWebView(addinId, devSettings.WebViewType.EdgeChromium);
        assert.strictEqual(await devSettingsWindows.getWebView(addinId), "Edge Chromium");
      });
      it("webView can be cleared", async function() {
        await devSettings.setWebView(addinId, devSettings.WebViewType.Default);
        assert.strictEqual(await devSettingsWindows.getWebView(addinId), undefined);
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
  }
});

describe("can specify to open dev tools", function() {
  if (isWindows) {
    before("debugging should be disabled", async function() {
      await devSettings.disableDebugging(addinId);
      const open = await devSettings.getOpenDevTools(addinId);
      assert.strictEqual(open, false);
    }),
    it("specify to open dev tools when debugging is enabled", async function() {
      await devSettings.enableDebugging(addinId, true, undefined, true);
      const open = await devSettings.getOpenDevTools(addinId);
      assert.strictEqual(open, true);
    });
    it("do not open dev tools by default when debugging is enabled", async function() {
      await devSettings.disableDebugging(addinId);
      const open = await devSettings.getOpenDevTools(addinId);
      assert.strictEqual(open, false);
    })
  }
});

describe("Appcontainer", async function() {
  if (isWindows) {
    describe("getAppcontainerName()", function() {
      it("developer add-in from https://localhost:3000", function() {
        assert.strictEqual(appcontainer.getAppcontainerName("https://localhost:3000/index.html"), "1_https___localhost_300004ACA5EC-D79A-43EA-AB47-E50E47DD96FC");
      });
      it("store add-in (ScriptLab)", function() {
        assert.strictEqual(appcontainer.getAppcontainerName("https://script-lab.azureedge.net", true), "0_https___script-lab.azureedge.net04ACA5EC-D79A-43EA-AB47-E50E47DD96FC");
      });
    });
    describe("getAppcontainerNameFromManifest()", function() {
      let sandbox: sinon.SinonSandbox;
      beforeEach(function() {
        sandbox = sinon.createSandbox();
      });
      afterEach(function() {
        sandbox.restore();
      });
      it("undefined source location", async function() {
        const manifest = {defaultSettings: ""};
        const readManifestFile = sinon.fake.returns(manifest);
        sandbox.stub(OfficeAddinManifest, "readManifestFile").callsFake(readManifestFile);
        try {
          await appcontainer.getAppcontainerNameFromManifest("https://localhost:3000/index.html");
          assert.strictEqual(0, 1); // expecting exception
        } catch (err: any) {
          assert.strictEqual(err.toString().includes("The source location could not be retrieved from the manifest."), true);
        }
      });
      it("valid source location", async function() {
        const sourceLocation = {sourceLocation: "https://localhost"};
        const manifest = {defaultSettings: sourceLocation};
        const readManifestFile = sinon.fake.returns(manifest);
        sandbox.stub(OfficeAddinManifest, "readManifestFile").callsFake(readManifestFile);
        const appcontainerName =  await appcontainer.getAppcontainerNameFromManifest("https://localhost");
        assert.strictEqual(appcontainerName, "1_https___localhost04ACA5EC-D79A-43EA-AB47-E50E47DD96FC");
      });
    });
  }
});

describe("Registration", function() {
  if (isWindows || isMac) {
    const manifestsFolder = fspath.resolve("test/files/manifests");

    this.beforeAll(async function() {
      await devSettings.unregisterAllAddIns();
    });
    describe("basic functionality", function() {
      it("No add-ins should be registered", async function() {
        const registered = await devSettings.getRegisterAddIns();
        assert.strictEqual(registered.length, 0);
      });
      it("Can register an add-in", async function() {
        const manifestPath = fspath.resolve(manifestsFolder, "manifest.xml");
        const registeredManifestPath = isMac
          ? fspath.join(os.homedir(), "/Library/Containers/com.microsoft.Excel/Data/Documents/wef/6dd581d2-98d1-4eaf-9506-e0a24be515f5.manifest.xml")
          : manifestPath;
        await devSettings.registerAddIn(manifestPath);
        const registeredAddins = await devSettings.getRegisterAddIns();
        const [registeredAddin] = registeredAddins;
        assert.strictEqual(registeredAddins.length, 1);
        assert.strictEqual(registeredAddin.id, "6dd581d2-98d1-4eaf-9506-e0a24be515f5");
        assert.strictEqual(registeredAddin.manifestPath, registeredManifestPath);
      });
      it("Can unregister an add-in", async function() {
        const manifestPath = fspath.resolve(manifestsFolder, "manifest.xml");
        await devSettings.unregisterAddIn(manifestPath);
        const registeredAddins = await devSettings.getRegisterAddIns();
        assert.strictEqual(registeredAddins.length, 0);
      });
    });
    describe("multiple add-ins", function() {
      const firstManifestPath = fspath.resolve(manifestsFolder, "manifest.xml");
      const secondManifestPath = fspath.resolve(manifestsFolder, "manifest2.xml");
      const firstManifestId = "6dd581d2-98d1-4eaf-9506-e0a24be515f5";
      const secondManifestId = "813cfc85-2a0f-49f6-8024-8d942cb73456";
      const firstRegisteredManifestPath = isMac
        ? fspath.join(os.homedir(), "/Library/Containers/com.microsoft.Excel/Data/Documents/wef/6dd581d2-98d1-4eaf-9506-e0a24be515f5.manifest.xml")
        : firstManifestPath;
      const secondRegisteredManifestPath = isMac
        ? fspath.join(os.homedir(), "/Library/Containers/com.microsoft.Excel/Data/Documents/wef/813cfc85-2a0f-49f6-8024-8d942cb73456.manifest2.xml")
        : secondManifestPath;

      it("Can register two add-ins", async function() {
        await devSettings.registerAddIn(firstManifestPath);
        await devSettings.registerAddIn(secondManifestPath);
        const registeredAddins = await devSettings.getRegisterAddIns();
        const [first, second] = registeredAddins;
        assert.strictEqual(registeredAddins.length, 2);
        assert.strictEqual(first.id, firstManifestId);
        assert.strictEqual(second.id, secondManifestId);
        assert.strictEqual(first.manifestPath, firstRegisteredManifestPath);
        assert.strictEqual(second.manifestPath, secondRegisteredManifestPath);
      });
      it("Can unregister one add-in", async function() {
        await devSettings.unregisterAddIn(secondManifestPath);
        const registeredAddins = await devSettings.getRegisterAddIns();
        const [first] = registeredAddins;
        assert.strictEqual(registeredAddins.length, 1);
        assert.strictEqual(first.id, firstManifestId);
        assert.strictEqual(first.manifestPath, firstRegisteredManifestPath);
      });
    });
    describe("json manifest functionality", function() {
      it('Error when file does not exist', () => {
        const zipPath = fspath.resolve(manifestsFolder, "nozip.zip");
        return registerWithTeams(zipPath)
          .then(() => {
              Promise.reject(new Error('Expected method to reject.'));
            },
            (error) => {
              assert.ok(error instanceof Error, "Should throw an error when zip doesn't exist");
              assert.strictEqual(error.message, `The file '${zipPath}' is not valid`);
            }
          );
      });
      it('Error when not a zip or xml file', () => {
        const zipPath = fspath.resolve(manifestsFolder, "json", "assets", "icon-16.png");
        return registerWithTeams(zipPath)
          .then(() => {
              Promise.reject(new Error('Expected method to reject.'));
            },
            (error) => {
              assert.ok(error instanceof Error, "Should throw an error when file is not a zip file");
              assert.strictEqual(error.message, `The file '${zipPath}' is not valid`);
            }
          );
      });
    });

    if (isWindows) {
      describe("register Outlook add-in", function() {
        this.beforeEach(async function() {
          const key = await devSettingsWindows.getDeveloperSettingsRegistryKey(devSettingsWindows.OutlookSideloadManifestPath);
          if (key) {
            await deleteKey(key);
          }
          await devSettings.unregisterAllAddIns();
        });
        this.afterEach(async function() {
          const key = await devSettingsWindows.getDeveloperSettingsRegistryKey(devSettingsWindows.OutlookSideloadManifestPath);
          if (key) {
            await deleteKey(key);
          }
          await devSettings.unregisterAllAddIns();
        });
        it("verify OutlookSideloadManifestPath key is not set and Develoer registry key is set", async function() {
          const manifestsFolder = fspath.resolve("test/files/manifests");
          const manifestPath = fspath.resolve(manifestsFolder, "manifest.xml");
          await devSettings.registerAddIn(manifestPath);

          // Verify manifest is set in Developer registry key
          const registeredAddins = await devSettings.getRegisterAddIns();
          const [registeredAddin] = registeredAddins;
          assert.strictEqual(registeredAddins.length, 1);
          assert.strictEqual(registeredAddin.id, "6dd581d2-98d1-4eaf-9506-e0a24be515f5");
          assert.strictEqual(registeredAddin.manifestPath, manifestPath);

          // Verify OutlookSideloadManifestPath is not set
          const key = devSettingsWindows.getDeveloperSettingsRegistryKey(devSettingsWindows.OutlookSideloadManifestPath);
          const value = await getStringValue(key, "");
          assert.strictEqual(value === undefined, true);
        });
      });
    }
  }
});

describe("RuntimeLogging", async function() {
  if (isWindows) {
    let pathBeforeTests: string | undefined;
    const testExecDirName = "testExec";
    const defaultFileName = "OfficeAddins.log.txt";
    const baseDirPath = process.cwd();
    const testExecDirPath = fspath.normalize(`${baseDirPath}/${testExecDirName}`);
    const defaultPath = fspath.normalize(`${process.env.TEMP}/${defaultFileName}`);

    this.beforeAll(async function() {
      pathBeforeTests = await devSettings.getRuntimeLoggingPath();
      await devSettings.disableRuntimeLogging();
    });

    this.afterAll(async function() {
      if (pathBeforeTests) {
        try {
          await devSettings.enableRuntimeLogging(pathBeforeTests);
        } catch (err) {
          console.log("Unable to restore original runtime logging settings. Runtime logging will be disabled.");
          await devSettings.disableRuntimeLogging();
        }
      } else {
        await devSettings.disableRuntimeLogging();
      }
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
        } catch (err: any) {
          error = err;
        }
        assert.ok(error instanceof Error, "should throw an error");
        assert.strictEqual(error.message, `You need to specify the path where the file can be written. Unable to write to: "${directoryDoesNotExist}\\${defaultFileName}".`);
        process.env.TEMP = TEMP;
      });
      it("default path but TEMP env var is not defined", async function() {
        const env = process.env;
        process.env = {};
        let error;
        try {
          const path: string = await devSettings.enableRuntimeLogging();
        } catch (err: any) {
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
          } catch (err: any) {
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
          } catch (err: any) {
            error = err;
          }
          assert.ok(error instanceof Error, "should throw an error");
          assert.strictEqual(error.message, `You need to specify the path to a writable file. Unable to write to: "${filePath}".`);
        });
      });
    });
  }
});

describe("Sideload to Desktop", function() {
  const manifestsFolder = fspath.resolve("test/files/manifests");
  const manifestPath = fspath.resolve(manifestsFolder, "manifest.xml");
  it("Verify pathToWrite not undefined", async function() {
    const manifest = await OfficeAddinManifest.readManifestFile(manifestPath);
    const pathToWrite: string = await devSettingsSideload.generateSideloadFile(OfficeApp.Excel, manifest);
    assert.notStrictEqual(pathToWrite, undefined);
  });
  it("Specify document - veriy pathToWrite not undefined", async function() {
    const manifest = await OfficeAddinManifest.readManifestFile(manifestPath);
    const document: string | undefined = devSettingsSideload.getTemplatePath(OfficeApp.Excel, AddInType.TaskPane);
    const pathToWrite: string = await devSettingsSideload.generateSideloadFile(OfficeApp.Excel, manifest, document);
    assert.notStrictEqual(pathToWrite, undefined);
  });
  it("Sideload unsupported host (expect error)'", async function() {
    let error;
    const manifestPath = fspath.resolve(manifestsFolder, "manifest.unsupportedhost.xml");
    try {
      await devSettingsSideload.sideloadAddIn(manifestPath, OfficeApp.Project, true /* canPrompt */, AppType.Desktop, undefined /* document */);
    } catch (err: any) {
      error = err;
    }
    assert.ok(error instanceof Error, "should throw an error");
    assert.strictEqual(error.message, "Sideload to the Project app is not supported.");
  })
});

describe("Sideload to web", function() {
  const docurl: string = "https://microsoft-my.sharepoint-df.com/personal/user_microsoft_com/_layouts/15/Doc.aspx?&file=Document.docx";
  const expectedQueryParams = "&wdaddindevserverport=8080&wdaddinmanifestfile=manifest.xml&wdaddinmanifestguid=6dd581d2-98d1-4eaf-9506-e0a24be515f5";
  let expectedUrl: string = `${docurl}${expectedQueryParams}`;
  const manifestsFolder = fspath.resolve("test/files/manifests");
  let manifestPath = fspath.resolve(manifestsFolder, "manifest.xml");
  it("Get sideload url with query params", async function() {
    const manifest: ManifestInfo = await OfficeAddinManifest.readManifestFile(manifestPath);
    const manifestFileName = fspath.basename(manifestPath)
    const generatedUrl: string | undefined = await devSettingsSideload.generateSideloadUrl(manifestFileName, manifest, docurl);
    assert.strictEqual(generatedUrl, expectedUrl);
  });
  it("Get sideload url with query params when isTest equals 'true'", async function() {
    const manifest: ManifestInfo = await OfficeAddinManifest.readManifestFile(manifestPath);
    const manifestFileName = fspath.basename(manifestPath)
    const generatedUrl: string | undefined = await devSettingsSideload.generateSideloadUrl(manifestFileName, manifest, docurl, true /* isTest */);
    const expectedTestQueryParam: string = "&wdaddintest=true";
    expectedUrl = `${expectedUrl}${expectedTestQueryParam}`;
    assert.strictEqual(generatedUrl, expectedUrl);
  });
  it("Sideload unsupported source location (expect error)'", async function() {
    let error;
    let manifestPath = fspath.resolve(manifestsFolder, "manifest.invalidsourcelocationforweb.xml");
    try {
      await devSettingsSideload.sideloadAddIn(manifestPath, OfficeApp.Excel, true /* canPrompt */, AppType.Web, docurl);
    } catch (err: any) {
      error = err;
    }
    assert.ok(error instanceof Error, "should throw an error");
    assert.strictEqual(error.message, "The hostname specified by the SourceLocation in the manifest is not supported for sideload. The hostname should be 'localhost' or 127.0.0.1.");
  });
  it("Sideload no document parameter provided (expect error)'", async function() {
    let error;
    let manifestPath = fspath.resolve(manifestsFolder, "manifest.xml");
    try {
      await devSettingsSideload.sideloadAddIn(manifestPath, OfficeApp.Excel, true /* canPrompt */, AppType.Web);
    } catch (err: any) {
      error = err;
    }
    assert.ok(error instanceof Error, "should throw an error");
    assert.strictEqual(error.message, `For sideload to web, you need to specify a document url.`);
  })
  it("Sideload unsupported Outlook host (expect error)'", async function() {
    let error;
    let manifestPath = fspath.resolve(manifestsFolder, "manifest.outlook.xml");
    try {
      await devSettingsSideload.sideloadAddIn(manifestPath, OfficeApp.Outlook, true /* canPrompt */, AppType.Web, docurl);
    } catch (err: any) {
      error = err;
    }
    assert.ok(error instanceof Error, "should throw an error");
    assert.strictEqual(error.message, `Sideload to the Outlook web app is not supported.`);
  })
});
