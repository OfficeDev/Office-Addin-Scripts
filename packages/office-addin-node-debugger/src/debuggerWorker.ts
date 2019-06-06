// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.
/*
 * Copyright (c) 2015-present, Facebook, Inc.
 *
 * This source code is licensed under the MIT license found in the
 * LICENSE file in the root directory of this source tree.
 *
 * @format
 */
/* global __fbBatchedBridge, self, importScripts, postMessage, onmessage: true */
/* eslint no-unused-vars: 0 */

import * as fetch from 'node-fetch';
declare var __fbBatchedBridge: any;
declare var __platformBundles: any; 

process.on('message', message => {
  let shouldQueueMessages = false;
  const messageQueue: { (): void }[] = [];

  const processEnqueuedMessages = () => {
    while (messageQueue.length) {
      const messageProcess = messageQueue.shift();
      if (messageProcess) {
        messageProcess();
      }
    }
    shouldQueueMessages = false;
  };

  const messageHandlers: any = {
    executeApplicationScript(message: any): void {
      for (const key in message.inject) {
        (global as any)[key] = JSON.parse(message.inject[key]);
      }

      shouldQueueMessages = true;

      function evalJS(js: string): void {
        try {
          eval(js.replace(/this\["webpackHotUpdate"\]/g, 'self["webpackHotUpdate"]').replace('GLOBAL', 'global'));
        } catch (error) {
          console.log(`Error Message: ${error.message}`);
          console.log(`Error stack: ${error.stack}`);
        } finally {
          if (process.send) {
            process.send({ replyID: message.id });
            processEnqueuedMessages();
          }
        }
      }

      // load platform bundles
      if ((global as any).__platformBundles != undefined) {
        const platformBundles = (global as any).__platformBundles.concat();
        delete (global as any).__platformBundles;   
        for (const [index, pb] of platformBundles.entries()) {
          //console.log(`PB start ${index + 1}/${platformBundles.length}`);
          eval(pb);
          //console.log(`PB done  ${index + 1}/${platformBundles.length}`);
        }
      }
      
      fetch
        .default(message.url)
        .then(resp => resp.text())
        .then(evalJS);
    }
  };

  const processMessage = () => {
    const sendReply = function(result: any, error?: any) {
      if (process.send) {
        process.send({ replyID: message.id, result, error });
      }
    };

    const handler = messageHandlers[message.method];

    // Special cased handlers
    if (handler) {
      handler(message);
      return;
    }

    // Other methods get called on the bridge
    let returnValue = [[], [], [], 0];
    try {
      if (typeof __fbBatchedBridge === 'object') {
        returnValue = __fbBatchedBridge[message.method].apply(null, message.arguments);
      }
    } finally {
      sendReply(JSON.stringify(returnValue));
    }
  };

  if (shouldQueueMessages) {
    messageQueue.push(processMessage);
  } else {
    processMessage();
  }
});
