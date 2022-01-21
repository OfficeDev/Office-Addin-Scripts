#!/usr/bin/env node

// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as commander from "commander";
import { logErrorMessage } from "office-addin-usage-data";
import * as commands from "./commands";

/* global process */

commander.name("office-addin-manifest");
commander.version(process.env.npm_package_version || "(version not available)");

commander.command("info <manifest-path>").action(commands.info);

commander
  .command("modify <manifest-path>")
  .option("-g,--guid [guid]", "Change the guid. A random guid is used if one is not provided.")
  .option("-d, --displayName <name>", "Change the display name.")
  .action(commands.modify);

commander
  .command("validate <manifest-path>")
  .option("-p, --production", "Verify the manifest for production environment")
  .action(commands.validate);

commander
  .command("export")
  .option("-m, --manifest <manfest-path>", "Specify the location of the manifest file.  Default is './manifest.json'")
  .option("-a, --assets <assets-folder-path>", "Specify the location of the assets folder.  Default is './assets'")
  .option(
    "-o, --output <output-path>",
    "Specify where to save the package.  Default is next to the manifest file input"
  )
  .action(commands.exportManifest);

// if the command is not known, display an error
commander.on("command:*", function () {
  logErrorMessage(`The command syntax is not valid.\n`);
  process.exitCode = 1;
  commander.help();
});

if (process.argv.length > 2) {
  commander.parse(process.argv);
} else {
  commander.help();
}
