// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as assert from "assert";
import * as mocha from "mocha";
import * as testHelper from "../../office-addin-test-helpers";
import * as defaults from "../src/defaults";
import { TestServer } from "../src/testServer";
const testServer = new TestServer();
const platformName = testServer.getPlatformName();
const promiseStartTestServer = testServer.startTestServer();
const testKey: string = "TestString";
const testValue: string = "Office-Addin-Test-Infrastructure";
const testValues: any = [];

describe("End-to-end validation of test server", function() {
    this.beforeAll(async function() {
        _setTLSRejectUnauthorized(false);
    });

    this.afterAll(async function() {
        _setTLSRejectUnauthorized(false);
    });
    describe("Setup test server", function() {
        it("Test server should have started", async function() {
            const startTestServer = await promiseStartTestServer;
            assert.equal(startTestServer, true);
        });
        it(`Http test server port should be ${defaults.httpPort}`, async function() {
            assert.equal(testServer.getTestHttpServerPort(), defaults.httpPort);
        });
        it(`Https test server port should be ${defaults.httpsPort}`, async function() {
            assert.equal(testServer.getTestHttpsServerPort(), defaults.httpsPort);
        });
        it(`Test server state should be set to true (i.e. started)`, async function() {
            assert.equal(testServer.getTestServerState(), true);
        });
    });
    describe("Ping server for response", function() {
        it("Http Test server should have responded to ping", async function() {
            const testServerResponse = await testHelper.pingTestServer(defaults.httpPort, false /* useHttps */);
            assert.equal(testServerResponse !== undefined, true);
            assert.equal(testServerResponse["status"], 200);
            assert.equal(testServerResponse["platform"], platformName);
        });
        it("Https Test server should have responded to ping", async function() {
            const testServerResponse = await testHelper.pingTestServer(defaults.httpsPort);
            assert.equal(testServerResponse !== undefined, true);
            assert.equal(testServerResponse["status"], 200);
            assert.equal(testServerResponse["platform"], platformName);
        });
    });
    describe("Send data to server and get results", function() {
        it("Send data should have succeeded", async function() {
            const sendData: boolean = await _sendTestData();
            assert.equal(sendData, true);
        });
        it("Getting sent data back from test server should succeed", async function() {
            const getResults: any = await testServer.getTestResults();
            assert.equal(getResults[0].Name, testKey);
            assert.equal(getResults[0].Value, testValue);
        });
    });
    describe("Stop test server", function() {
        it("Test server should have stopped ", async function() {
            const stopTestServer: boolean = await testServer.stopTestServer();
            assert.equal(stopTestServer, true);
        });
        it(`Dev-server state should be set to false (i.e. stopped)`, async function() {
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
    return testHelper.sendTestResults(testValues, defaults.httpsPort);
}

function _setTLSRejectUnauthorized(on: boolean) {
    process.env.NODE_TLS_REJECT_UNAUTHORIZED = on ? " 1" : "0";
}
