#!/usr/bin/env node

// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as commander from "commander";
import * as commands from "./commands";

commander
.command("info <manifest-path>")
.action(commands.info);

commander
.command("modify <manifest-path>")
.option("-g,--guid [guid]", "Change the guid. A random guid is used if one is not provided.")
.option("-d, --displayName <name>", "Change the display name.")
.action(commands.modify);

commander.parse(process.argv);
