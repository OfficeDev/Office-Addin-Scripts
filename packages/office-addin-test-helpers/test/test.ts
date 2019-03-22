import * as assert from "assert";
import * as mocha from "mocha";
import * as testHelper from "../src/testHelpers";
import { TestServer } from "../../office-addin-test-server";
const port: number = 4201;
const testServer = new TestServer(port);
const promiseStartTestServer = testServer.startTestServer(true /* mochaTest */);
const testKey: string = "TestString";
const testValue: string = "Office-Addin-Test-Infrastructure";
const testValues: any = [];

describe("Start test server, validate pingTestServer and sendTestResults methods and stop test server", function () {
    describe("Setup test server", function () {
        it("Test server should have started", async function () {
            const startTestServer = await promiseStartTestServer;
            assert.equal(startTestServer, true);
        });
        it("Test server should have responded to ping", async function () {
            const testServerResponse: object = await testHelper.pingTestServer(port);
            assert.equal(testServerResponse["status"], 200);
            assert.equal(testServerResponse["platform"], testServer.getPlatformName());
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