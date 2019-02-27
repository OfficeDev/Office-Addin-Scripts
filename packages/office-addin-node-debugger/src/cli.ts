#!/usr/bin/env node

// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.
// Copyright (c) 2015-present, Facebook, Inc.
// This source code is licensed under the MIT license found in the
// LICENSE file in the root directory of this source tree.

import * as commander from 'commander';
import { run } from './debugger';

commander
.option('-h, --host <host>', 'The hostname where the packager is running.')
.option('-p, --port <port>', 'The port where the packager is running.')  
.parse(process.argv);

run(commander.host, commander.port);


