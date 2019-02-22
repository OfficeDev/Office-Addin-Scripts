import * as assert from "assert";
import * as testServer from "../src/testServer";
const port: string = "8080";
const testKey: string = "TestString";
const testValue: string = "Office-Addin-Test-Infrastructure";
const testValues: any = [];

describe("End-to-end validation of test server", function() {
    describe("1) Setup test server, 2) ping server for response, 3) send data to server and get results, 4) stop test server", function() {
        it("Dev-server should have started", async function() {
            const startTestServer = await testServer.startTestServer(port);
            assert.equal(startTestServer, true);
            assert.equal(testServer.getTestServerState(), true);
            assert.equal(testServer.getTestServerPort(), port);
        });
        it("Test server should have responded to ping", async function() {
            const testServerResponse: Object = await testServer.pingTestServer();
            assert.equal(testServerResponse["status"], 200);
            assert.equal(testServerResponse["platform"], process.platform === "win32" ? "Win32" : "Mac");
        });
        it("Test server should have received sent data and getTestResults should have retrieved data ", async function() {
            const sendData = await _sendTestData();
            const getResults: JSON = testServer.getTestResults();
            assert.equal(sendData, true);
            assert.equal(getResults[0].Name, testKey);
            assert.equal(getResults[0].Value, testValue);
        });
        it("Test server should have stopped ", async function () {
            const stopTestServer = await testServer.stopTestServer();
            assert.equal(stopTestServer, true);
            assert.equal(testServer.getTestServerState(), false);
        });
    });
});

async function _sendTestData(): Promise<boolean> {
    const testData: Object = {};
    const nameKey = "Name";
    const valueKey = "Value";

    testData[nameKey] = testKey;
    testData[valueKey] = testValue;
    testValues.push(testData);

    return testServer.sendTestResults(testValues);
}
