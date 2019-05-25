// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.
//
// The intension of the file is to avoid error during "npm install" in the root directory.
// package.json bin field calls into this file first. If ./lib/cli.js is used directly,
// lerna bootstrap will cause an error.
require("./lib/cli.js");