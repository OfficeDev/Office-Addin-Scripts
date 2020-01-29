// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as commander from "commander";
import {logErrorMessage, parseNumber} from "office-addin-cli";
import * as defaults from "./defaults";
import {ensureCertificatesAreInstalled} from "./install";
import {deleteCertificateFiles, uninstallCaCertificate} from "./uninstall";
import {verifyCertificates} from "./verify";
import * as usageDataHelper from "./usagedata-helper";

function parseDays(optionValue: any): number | undefined {
    const days = parseNumber(optionValue, "--days should specify a number.");

    if (days !== undefined) {
        if (!Number.isInteger(days)) {
            throw new Error("--days should be integer.");
        }
        if (days <= 0) {
            throw new Error("--days should be greater than zero.");
        }
    }
    return days;
}

export async function install(command: commander.Command) {
    try {
        const days = parseDays(command.days);

        await ensureCertificatesAreInstalled(days, command.machine).then(() => {
            usageDataHelper.sendUsageDataSuccessEvent("install");
        });
    } catch (err) {
        usageDataHelper.sendUsageDataException("install", `${err}`);
        logErrorMessage(err);
    }
}

export async function verify(command: commander.Command) {
    try {
        if (await verifyCertificates()) {
            console.log(`You have trusted access to https://localhost.\nCertificate: ${defaults.localhostCertificatePath}\nKey: ${defaults.localhostKeyPath}`);
            usageDataHelper.sendUsageDataSuccessEvent("verify");
        } else {
            console.log(`You need to install certificates for trusted access to https://localhost.`);
        }
    } catch (err) {
        logErrorMessage(err);
        usageDataHelper.sendUsageDataException("verify", `${err}`);
    }
}

export async function uninstall(command: commander.Command) {
    try {
        await uninstallCaCertificate(command.machine);
        await deleteCertificateFiles(defaults.certificateDirectory);
        usageDataHelper.sendUsageDataSuccessEvent("uninstall");
    } catch (err) {
        usageDataHelper.sendUsageDataException("uninstall", `${err}`);
        logErrorMessage(err);
    }
}
