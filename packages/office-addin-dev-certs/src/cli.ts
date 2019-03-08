#!/usr/bin/env node

// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as commander from "commander";
import * as commands from "./commands";
import {defaultCaCertPath, defaultCertPath, defaultKeyPath} from "./default";

commander
    .command("generate")
    .option("--ca-cert <ca-cert-path>", `Specifies the path where the CA certificate file is written. Default: ${defaultCaCertPath}`)
    .option("--cert <cert-path>", `Specifies the path where the SSL certificate file is written. Default: ${defaultCertPath}`)
    .option("--key <key-path>", `Specifies the path where the private key for the SSL certificate file is written. Default: ${defaultKeyPath}`)
    .option("--install", `Install the generated CA certificate`)
    .description(`Generate an SSL certificate for localhost and a CA certificate which has issued it.`)
    .action(commands.generate);

commander
    .command("install <ca-cert-path>")
    .description(`Install the CA certificate.`)
    .action(commands.install);

commander
    .command("verify")
    .description(`Verify the CA certificate.`)
    .action(commands.verify);

commander
    .command("uninstall")
    .description(`Uninstall the certificate.`)
    .action(commands.uninstall);

commander.parse(process.argv);
