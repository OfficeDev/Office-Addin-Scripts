#!/usr/bin/env node

// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.
//
// If the package.json bin config specifies a file in the lib folder, it will cause an
// error during "npm install" if the lib folder doesn't exist (because the package hasn't been built yet).
// It specifies this file instead which then calls into the file in the lib folder.
require("./lib/cli.js");
