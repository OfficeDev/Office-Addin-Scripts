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

describe("Start test server, ping test server send test results and stop test server", function () {
    before("Test Server should have started and responded to ping", async function () {
        const startTestServer = await promiseStartTestServer;
        assert.equal(startTestServer, true);
        const serverResponse = await testHelper.pingTestServer(port);
        assert.equal(serverResponse["status"], 200);
    });
    describe("Send data to test server", function () {
        it("Send data should have succeeded", async function () {
            const sendData: boolean = await _sendTestData();
            assert.equal(sendData, true);
        });
    });
    after("Test server should have stopped server state should be set to false", async function () {
        const stopTestServer: boolean = await testServer.stopTestServer();
        assert.equal(stopTestServer, true);
        assert.equal(testServer.getTestServerState(), false);
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