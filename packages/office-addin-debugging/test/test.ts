// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as assert from "assert";
import express from "express";
import * as fs from "fs";
import * as http from "http";
import * as mocha from "mocha";
import * as debugInfo from "../src/debugInfo";
import * as port from "../src/port";
import * as start from "../src/start";
import * as stop from "../src/stop";

/* global console */

function startServer(serverPort: number): http.Server {
  const server = http.createServer(express());

  server.on("close", () => {
    console.log(`Server has stopped listening on port ${serverPort}.`);
  });

  server.on("listening", () => {
    console.log(`Server is listening on port ${serverPort}.`);
  });

  server.listen(serverPort);

  return server;
}

describe("port functions", function() {
  describe("getProcessIdsForPort()", async function() {
    let portNotInUse: number;
    let serverPort: number;
    let server: http.Server;

    before(async function() {
      serverPort = await port.randomPortNotInUse();
      server = startServer(serverPort);
      portNotInUse = await port.randomPortNotInUse();
      console.log(`Port ${portNotInUse} is not in use.`);
    });
    it("no process ids", async function() {
      const processIds = await port.getProcessIdsForPort(portNotInUse);
      assert.strictEqual(Array.isArray(processIds), true);
      assert.strictEqual(processIds.length, 0);
    });
    it("one process id", async function() {
      const processIds = await port.getProcessIdsForPort(serverPort);
      assert.strictEqual(Array.isArray(processIds), true);
      assert.strictEqual(processIds.length, 1);
      assert.strictEqual(processIds[0], process.pid);
    });
    after(function() {
      server.close();
    });
  });

  describe("isPortInUse()", async function() {
    let portNotInUse: number;
    let serverPort: number;
    let server: http.Server;

    before(async function() {
      serverPort = await port.randomPortNotInUse();
      server = startServer(serverPort);
      portNotInUse = await port.randomPortNotInUse();
      console.log(`Port ${portNotInUse} is not in use.`);
    });
    it("port not in use", async function() {
      assert.strictEqual(await port.isPortInUse(portNotInUse), false);
    });
    it("port is in use", async function() {
      assert.strictEqual(await port.isPortInUse(serverPort), true);
    });
    it("port is no longer in use", async function() {
      server.close();
      // verify the port is no longer in use
      assert.strictEqual(await port.isPortInUse(serverPort), false);
    });
  });

});

describe("start/stop functions", function() {
  const pid = 1234;
  it("writing process id file", async function() {
    await debugInfo.saveDevServerProcessId(pid);
    const json = fs.readFileSync(debugInfo.getDebuggingInfoPath());
    const devServerInfo = JSON.parse(json.toString());
    const processId = devServerInfo.devServer.processId;
    assert.strictEqual(processId.toString(), pid.toString());
  });
  it("reading process id file", async function() {
    const id = debugInfo.readDevServerProcessId();
    if (id) {
      assert.strictEqual(id.toString(), pid.toString());
    }
  });
  it("deleting process id file", async function() {
    debugInfo.clearDevServerProcessId();
  });
  it("read process id file that is missing", async function() {
    const id = debugInfo.readDevServerProcessId();
    assert.strictEqual(id, undefined);
  });
  it("write process id file that already exists", async function() {
    const secondPid = 5678;
    await debugInfo.saveDevServerProcessId(pid);
    await debugInfo.saveDevServerProcessId(secondPid);
    const json = fs.readFileSync(debugInfo.getDebuggingInfoPath());
    const devServerInfo = JSON.parse(json.toString());
    const processId = devServerInfo.devServer.processId;
    assert.strictEqual(processId, secondPid);
    debugInfo.clearDevServerProcessId();
  });
  it("read process id file with corrupt data", async function() {
    let badIdValue;
    const corruptId = '{"devServer":{"processId":"bad id"}}';
    const errorMessageInvalidProcessId = "Invalid process id";
    fs.writeFileSync(debugInfo.getDebuggingInfoPath(), corruptId);
    try {
      badIdValue = debugInfo.readDevServerProcessId();
    } catch (error: any) {
      assert.ok(error.message.includes(errorMessageInvalidProcessId), "invalid process id error message found");
    }
    assert.strictEqual(badIdValue, undefined);
    debugInfo.clearDevServerProcessId();
    });
});
