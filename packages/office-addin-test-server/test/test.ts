import * as assert from "assert";
import * as mocha from "mocha";
import * as testServer from "../src/testServer";
const port: number = 8080;
const testKey: string = "TestString";
const testValue: string = "Office-Addin-Test-Infrastructure";
const testValues: any = [];

describe("End-to-end validation of test server", function() {
    describe("Setup test server", function() {
        it("Dev-server should have started", async function() {
            const startTestServer = await testServer.startTestServer(port);
            assert.equal(startTestServer, true);
        });
        it(`Dev-server port should be ${port}`, async function () {
            assert.equal(testServer.getTestServerPort(), port);
        });
        it(`Dev-server state should be set to true (i.e. started)`, async function () {
            assert.equal(testServer.getTestServerState(), true);
        });
    });

    describe("Ping server for response", function () {
        let testServerResponse: any;
        it("Test server should have responded to ping", async function () {
            testServerResponse = await testServer.pingTestServer(port);
            assert.equal(testServerResponse != undefined, true);
        });
        it("Test server status should be '200'", async function () {
            assert.equal(testServerResponse["status"], 200);
            assert.equal(testServerResponse["platform"], process.platform === "win32" ? "Win32" : "Mac");
        });
        it("Test server platform should reflect correct operating system", async function () {
            assert.equal(testServerResponse["platform"], process.platform === "win32" ? "Win32" : "Mac");
        });
    });
    describe("Send data to server and get results", function () {
        it("Send data should have succeeded", async function () {
            const sendData: boolean = await _sendTestData();
            assert.equal(sendData, true);
        });
        it("Getting sent data back from test server should succeed", async function () {
            const getResults: any = testServer.getTestResults();
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

    return testServer.sendTestResults(testValues, port);
}
