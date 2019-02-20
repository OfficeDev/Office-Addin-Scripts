import * as assert from "assert";
import officeAddinTestInfrastructure from "../src/office-addin-test-infrastructure";
const port: number = 8080;
const officeAddinTestInfra = new officeAddinTestInfrastructure(port);
const testKey: string = "TestString";
const testValue: string = "Office-Addin-Test-Infrastructure";
const testValues: any = [];
const XMLHttpRequest = require("xmlhttprequest").XMLHttpRequest;


describe("1) Setup test server, 2) ping server for response, 3) send data to server and get results, 4) stop test server", function() {
    describe("Start sideload, start dev-server, and start test-server", function() {
        it("Dev-server should have started", async function() {
            const startTestServer = await officeAddinTestInfra.startTestServer();
            assert.equal(startTestServer, true);
            assert.equal(officeAddinTestInfra.getTestServerState(), true);
            assert.equal(officeAddinTestInfra.getTestServerPort(), port);
        });
        it("Test server should have responded to ping", async function() {
            const testServerResponse = await _pingTestServer();
            assert.equal(testServerResponse, process.platform === "win32" ? "Win32" : "Mac");
        });
        it("Test server should have received sent data and getTestResults should have retrieved data ", async function() {
            _sendTestData();
            const testResults: any = await officeAddinTestInfra.getTestResults();
            assert.equal(testResults[0].Name, testKey);
            assert.equal(testResults[0].Value, testValue);
        });
        it("Test server should have stopped ", async function () {
            const stopTestServer = await officeAddinTestInfra.stopTestServer();
            assert.equal(stopTestServer, true);
            assert.equal(officeAddinTestInfra.getTestServerState(), false);
        });
    });
});

async function _pingTestServer(): Promise<string> {
    return new Promise<string>(async function(resolve) {
        const xhr = new XMLHttpRequest();
        const pingUrl: string = `https://localhost:${port}/ping`;
        xhr.onreadystatechange = (e) => {
            if (xhr.readyState === 4 && xhr.status === 200) {
                resolve(xhr.responseText);
            }
        };
        xhr.open("GET", pingUrl, true);
        xhr.send();
    });
}

async function _sendTestData(): Promise<void> {
    const testData: any = {};
    const nameKey = "Name";
    const valueKey = "Value";
    testData[nameKey] = testKey;
    testData[valueKey] = testValue;
    testValues.push(testData);

    const json = JSON.stringify(testValues);
    const xhr = new XMLHttpRequest();
    const url: string = `https://localhost:${port}/results/`;
    const dataUrl: string = url + "?data=" + encodeURIComponent(json);

    xhr.open("GET", dataUrl, true);
    xhr.send();
}
