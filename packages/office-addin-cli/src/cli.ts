#!/usr/bin/env node

// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as commander from "commander";

commander.name("office-addin-cli");
commander.version(process.env.npm_package_version || "(version not available)");

if (process.argv.length > 2) {
    commander.parse(process.argv);
} else {
    commander.help();
}
