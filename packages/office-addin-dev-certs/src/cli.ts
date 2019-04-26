#!/usr/bin/env node

// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as commander from "commander";
import * as commands from "./commands";
import * as defaults from "./defaults";

commander
    .command("install")
    .option("--machine", "Specifies to install the CA certificate to the machine root instead of the user root. Need to be administrator to use this option.")
    .option("--days <days>", `Specifies the validity of CA certificate in days. Default: ${defaults.daysUntilCertificateExpires}`)
    .description(`Generate an SSL certificate for "localhost" issued by a CA certificate which is installed.`)
    .action(commands.install);

commander
    .command("verify")
    .description(`Verify the CA certificate.`)
    .action(commands.verify);

commander
    .command("uninstall")
    .option("--machine", "Specifies to uninstall the CA certificate from the machine root instead of the user root. Need to be administrator to use this option.")
    .description(`Uninstall the certificate.`)
    .action(commands.uninstall);

commander.parse(process.argv);
