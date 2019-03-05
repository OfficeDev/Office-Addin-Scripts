#!/usr/bin/env node

// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as commander from "commander";
import * as commands from "./commands";

commander
    .command("generate")
    .option("--ca-cert ", "Specify the path where the CA certificate file is written.")
    .option("--cert ", "Specify the path where the SSL certificate file is written.")
    .option("--key ", "Specify the path where the private key for the SSL certificate file is written.")
    .action(commands.generate);

commander
    .command("install")
    .action(commands.install);

commander
    .command("verify")
    .action(commands.verify);

commander
    .command("uninstall")
    .action(commands.uninstall);

commander
    .command("clean")
    .action(commands.clean);

commander.parse(process.argv);
