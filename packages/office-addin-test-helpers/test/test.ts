// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as assert from "assert";
import * as mocha from "mocha";
import * as testHelper from "../src/testHelpers";
import { defaultHttpPort, defaultHttpsPort, TestServer } from "../../office-addin-test-server";
const port: number = 4201;
const testServer = new TestServer(port);
const promiseStartTestServer = testServer.startTestServer();
const testKey: string = "TestString";
const testValue: string = "Office-Addin-Test-Infrastructure";
const testValues: any = [];

describe("Start test server, validate pingTestServer and sendTestResults methods and stop test server", function () {
    describe("Setup test server", function () {
        it("Test server should have started", async function () {
            testHelper.setTLSRejectUnauthorized(false);
            const startTestServer = await promiseStartTestServer;
            assert.equal(startTestServer, true);
        });
        it("Test server should have responded to ping", async function () {
            // ping test server on https port
            const httpsTestServerResponse: object = await testHelper.pingTestServer(defaultHttpsPort);
            assert.equal(httpsTestServerResponse["status"], 200);
            assert.equal(httpsTestServerResponse["platform"], testServer.getPlatformName());
            // ping test server on http port
            const httpTestServerResponse: object = await testHelper.pingTestServer(defaultHttpPort, false /* https */);
            assert.equal(httpTestServerResponse["status"], 200);
            assert.equal(httpTestServerResponse["platform"], testServer.getPlatformName());
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

    return testHelper.sendTestResults(testValues, defaultHttpsPort);
}