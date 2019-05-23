// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as commander from "commander";
import * as commands from "./commands";
import { defaultPort } from "./testServer";

commander
    .command("start")
    .option(`-p --port [port number]", "Port number must be between 0 - 65535. Default: ${defaultPort}`)
    .action(commands.start);

commander.parse(process.argv);