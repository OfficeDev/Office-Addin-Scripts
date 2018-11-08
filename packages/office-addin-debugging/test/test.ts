import * as assert from "assert";
import * as express from "express";
import * as http from "http";
import * as mocha from "mocha";
import * as net from "net";
import * as ws from "ws";
import * as port from "../src/port";

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
