// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as assert from "assert";
import { describe, it } from "mocha";
import * as testHelper from "../src/testHelpers";
import { TestServer } from "office-addin-test-server";
const port: number = 4201;
const testServer = new TestServer(port);
const promiseStartTestServer = testServer.startTestServer(true /* mochaTest */);
const testKey: string = "TestString";
const testValue: string = "Office-Addin-Test-Infrastructure";
const testValues: any = [];

describe("Start test server, validate pingTestServer and sendTestResults methods and stop test server", function () {
  describe("Setup test server", function () {
    it("Test server should have started", async function () {
      // give the server some time to start
      this.timeout(10000);
      const startTestServer = await promiseStartTestServer;
      assert.equal(startTestServer, true);
    });
    it("Test server should have responded to ping", async function () {
      const testServerResponse: testHelper.TestServerResponse = await testHelper.pingTestServer(port);
      assert.equal(testServerResponse.status, 200);
      assert.equal(testServerResponse.platform, testServer.getPlatformName());
    });
    it("Send data should have succeeded", async function () {
      const sendData: boolean = await _sendTestData();
      assert.equal(sendData, true);
    });
    it("Test server should have stopped ", async function () {
      const stopTestServer: boolean = await testServer.stopTestServer();
      assert.equal(stopTestServer, true);
    });
  });
});

async function _sendTestData(): Promise<boolean> {
  const testData: any = {};
  const nameKey = "Name";
  const valueKey = "Value";

  testData[nameKey] = testKey;
  testData[valueKey] = testValue;
  testValues.push(testData);

  return testHelper.sendTestResults(testValues, port);
}
