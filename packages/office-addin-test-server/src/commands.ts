// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as commnder from "commander";
import { parseNumber } from "office-addin-cli";
import { defaultPort, TestServer } from "./testServer";

/* global console */

export async function start(command: commnder.Command) {
  const testServerPort: number = command.port !== undefined ? parseTestServerPort(command.port) : defaultPort;
  const testServer = new TestServer(testServerPort);
  const serverStarted: boolean = await testServer.startTestServer();

  if (serverStarted) {
    console.log(`Server started successfully on port ${testServerPort}`);
  } else {
    console.log("Server failed to start");
  }
}

function parseTestServerPort(optionValue: any): number {
  const testServerPort = parseNumber(optionValue, "--dev-server-port should specify a number.");

  if (testServerPort !== undefined) {
    if (testServerPort < 0 || testServerPort > 65535) {
      throw new Error("port should be between 0 and 65535.");
    }
  }
  return testServerPort;
}
