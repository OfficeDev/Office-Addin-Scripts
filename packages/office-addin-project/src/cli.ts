#!/usr/bin/env node

// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as commander from "commander";
import { logErrorMessage } from "office-addin-usage-data";
import * as commands from "./commands";

/* global process */

commander.name("office-addin-project");
commander.version(process.env.npm_package_version || "(version not available)");

// if the command is not known, display an error
commander.on("command:*", function () {
  logErrorMessage(`The command syntax is not valid.\n`);
  process.exitCode = 1;
  commander.help();
});

commander
  .command("convert")
  .option(
    "-m, --manifest <manifest-path>",
    "Specify the location of the manifest file.  Default is './manifest.xml'"
  )
  .option(
    "-b, --backup <backup-path>",
    "Specify the location of the backup folder for the project.  Default is './backup.zip'"
  )
  .option(
    "-p, --project <project-path>",
    "Specify the location of the root directory of the project.  Default is the directory of the manifest file."
  )
  .option(
    "--preview",
    "Use the devPreview version of the json schema in the manifest output"
  )
  .option( "--confirm", "Confirmes the conversion")
  .action(commands.convert);

if (process.argv.length > 2) {
  commander.parse(process.argv);
} else {
  commander.help();
}
