#!/usr/bin/env node

// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.
// Copyright (c) 2015-present, Facebook, Inc.
// This source code is licensed under the MIT license found in the
// LICENSE file in the root directory of this source tree.

import * as child from 'child_process';
import * as commander from 'commander';
import { fork } from 'child_process';
import WebSocket = require('ws');

commander
  .option('-h, --hostname <hostname>', 'The hostname where the packager is running.')
  .option('-p, --port <port>', 'The port where the packager is running.')
  .parse(process.argv);

const hostName: string = commander.hostname ? commander.hostname : 'localhost';
const port: string = commander.port ? commander.port : '8081';
const role: string = 'debugger';
const debuggerName: string = 'OfficeAddinDebugger';
const debuggerWorkerRelativePath: string = '\\debuggerWorker.js';
const debuggerWorkerFullPath: string = `${__dirname}${debuggerWorkerRelativePath}`;
const websocketRetryTimeout: number = 500;

(function() {
  function connectToDebuggerProxy(): void {
    var ws = new WebSocket(`ws://${hostName}:${port}/debugger-proxy?role=${role}&name=${debuggerName}`);
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
        console.clear();

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
})();
