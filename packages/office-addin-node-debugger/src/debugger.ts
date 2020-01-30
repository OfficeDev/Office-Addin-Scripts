#!/usr/bin/env node

// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.
// Copyright (c) 2015-present, Facebook, Inc.
// This source code is licensed under the MIT license found in the
// LICENSE file in the root directory of this source tree.

import * as child from 'child_process';
import * as commander from 'commander';
import { fork } from 'child_process';
import * as path from "path";
import WebSocket = require('ws');

export function run(host: string = "localhost", port: string = "8081", 
  role: string = "debugger", debuggerName: string = "OfficeAddinDebugger") {
    
  const debuggerWorkerRelativePath: string = '\\debuggerWorker.js';
  const debuggerWorkerFullPath: string = `${__dirname}${debuggerWorkerRelativePath}`;
  const websocketRetryTimeout: number = 500;

  function connectToDebuggerProxy(): void {
    var ws = new WebSocket(`ws://${host}:${port}/debugger-proxy?role=${role}&name=${debuggerName}`);
    var worker: child.ChildProcess;

    function createJSRuntime(): void {
      
      // This worker will run the application javascript code.
      worker = fork(`${debuggerWorkerFullPath}`, [], {
        stdio: ['pipe', 'pipe', 'pipe', 'ipc'],
        execArgv: ['--inspect']
      });
      worker.on('message', message => {
        ws.send(JSON.stringify(message));
      });
    }

    function shutdownJSRuntime(): void {
      if (worker) {
        worker.kill();
        worker.unref();
      }
    }

    ws.onopen = () => {
      console.log('Web socket opened...');
    };
    ws.onmessage = message => {
      if (!message.data) {
        return;
      }

      var object = JSON.parse(message.data.toString());

      if (object.$event === 'client-disconnected') {
        shutdownJSRuntime();
        return;
      }
      if (!object.method) {
        return;
      }
      // Special message that asks for a new JS runtime
      if (object.method === 'prepareJSRuntime') {
        shutdownJSRuntime();
        //console.clear();

        createJSRuntime();
        ws.send(JSON.stringify({ replyID: object.id }));
      } else if (object.method === '$disconnected') {
        shutdownJSRuntime();
      } else {
        worker.send(object);
      }
    };
    ws.onclose = e => {
      shutdownJSRuntime();
      if (e.reason) {
        console.log(`Web socket closed because the following reason: ${e.reason}`);
      }
      setTimeout(connectToDebuggerProxy, websocketRetryTimeout);
    };
    ws.onerror = event => {
      console.log(`${event.error}`);
    };
  }
  connectToDebuggerProxy();
}
