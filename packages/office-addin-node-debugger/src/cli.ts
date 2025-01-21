#!/usr/bin/env node

// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.
// Copyright (c) 2015-present, Facebook, Inc.
// This source code is licensed under the MIT license found in the
// LICENSE file in the root directory of this source tree.

import { Command } from "commander";
import { run } from "./debugger";

const commander = new Command();

commander
  .option("-h, --host <host>", "The hostname where the packager is running.")
  .option("-p, --port <port>", "The port where the packager is running.")
  .parse(process.argv);

const options = commander.opts();
run(options.host, options.port);
