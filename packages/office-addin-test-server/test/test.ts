import * as assert from "assert";
import * as mocha from "mocha";
import * as testHelper from "../../office-addin-test-helpers"
import { TestServer } from "../src/testServer";
const port: number = 4201;
const testServer = new TestServer(port);
const platformName = testServer.getPlatformName();
const promiseStartTestServer = testServer.startTestServer(true /* mochaTest */);
const testKey: string = "TestString";
const testValue: string = "Office-Addin-Test-Infrastructure";
const testValues: any = [];

describe("End-to-end validation of test server", function() {
    before("Test Server should have started and responded to ping", async function() {
            const startTestServer = await promiseStartTestServer;
            assert.equal(startTestServer, true);
            const serverResponse = await testHelper.pingTestServer(port);
            assert.equal(serverResponse["status"], 200);
            assert.equal(serverResponse["platform"], platformName);
    });
    describe("Get test server properties", function() {
        it(`Test server port should be ${port}`, async function () {
            assert.equal(testServer.getTestServerPort(), port);
        });
        it(`Test server state should be set to true (i.e. started)`, async function () {
            assert.equal(testServer.getTestServerState(), true);
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
    after("Test server should have stopped and server state should be set to false", async function () {
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
