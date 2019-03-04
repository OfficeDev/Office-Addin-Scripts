#!/usr/bin/env node

// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as commander from "commander";
import * as commands from "./commands";

commander
    .command("generate <manifest-path>")
    .option("--path <path>", "Specify the path of generated certficate files.")
    .action(commands.generate);

commander
    .command("install <manifest-path>")
    .action(commands.install);

commander
    .command("verify <manifest-path>")
    .action(commands.verify);

commander
    .command("uninstall")
    .action(commands.uninstall);

commander
    .command("clean <manifest-path>")
    .action(commands.clean);

commander.parse(process.argv);
