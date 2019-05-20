#!/usr/bin/env node

// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as commander from "commander";
import * as commands from "./commands";
import * as defaults from "./defaults";

commander.name("office-addin-dev-certs");
commander.version(process.env.npm_package_version || "(version not available)");


commander
    .command("install")
    .option("--machine", "Install the CA certificate for all users. You must be an Administrator.")
    .option("--days <days>", `Specifies the validity of CA certificate in days. Default: ${defaults.daysUntilCertificateExpires}`)
    .description(`Generate an SSL certificate for "localhost" issued by a CA certificate which is installed.`)
    .action(commands.install);

commander
    .command("verify")
    .description(`Verify the CA certificate.`)
    .action(commands.verify);

commander
    .command("uninstall")
    .option("--machine", "Uninstall the CA certificate for all users. You must be an Administrator.")
    .description(`Uninstall the certificate.`)
    .action(commands.uninstall);

if (process.argv.length > 2) {
    commander.parse(process.argv);
} else {
    commander.help();
}
