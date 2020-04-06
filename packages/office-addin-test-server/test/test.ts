// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as assert from "assert";
import * as mocha from "mocha";
import * as testHelper from "../../office-addin-test-helpers"
import { TestServer } from "../src/testServer";
const port: number = 4201;
const testKey: string = "TestString";
const testValue: string = "Office-Addin-Test-Infrastructure";
const testValues: any = [];
let testServer: TestServer;

describe("End-to-end validation of test server", function() {
    before(function() {
        // skip tests on Linux -- need to get dev certs working there
        if (process.platform == "linux") {
            this.skip();
        }
    })
    describe("Start test server", function() {
        it("Test server should have started", async function() {
            testServer = new TestServer(port);
            const started: boolean = await testServer.startTestServer(true /* mochaTest */);
            assert.equal(started, true, "Test server should have started.");
            assert.equal(testServer.getTestServerPort(), port, `Test server port should be ${port}`);
            assert.equal(testServer.getTestServerState(), true, "Test server state should true (started).");
        });
    });

    describe("Ping server for response", function () {
        let testServerResponse: any;
        it("Test server should have responded to ping", async function () {
            const platformName = testServer.getPlatformName();
            testServerResponse = await testHelper.pingTestServer(port);
            assert.equal(testServerResponse !== undefined, true);
            assert.equal(testServerResponse.status, 200);
            assert.equal(testServerResponse.platform, platformName);
        });
    });
    describe("Send data to server and get results", function () {
        it("Send data should have succeeded", async function () {
            const sendData: boolean = await _sendTestData();
            assert.equal(sendData, true);
        });
        it("Getting sent data back from test server should succeed", async function () {
            const getResults: any = await testServer.getTestResults();
            assert.equal(getResults[0].Name, testKey);
            assert.equal(getResults[0].Value, testValue);
        });
    });

    describe("Stop test server", function () {
        it("Test server should have stopped ", async function () {
            const stopTestServer: boolean = await testServer.stopTestServer();
            assert.equal(stopTestServer, true);
        });
        it(`Dev-server state should be set to false (i.e. stopped)`, async function () {
            assert.equal(testServer.getTestServerState(), false);
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
