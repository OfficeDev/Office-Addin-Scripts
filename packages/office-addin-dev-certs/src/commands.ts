import * as commander from "commander";
import {logErrorMessage, parseNumber} from "office-addin-cli";
import {generateCertificates} from "./generate";
import {installCaCertificate} from "./install";
import {uninstallCaCertificate} from "./uninstall";
import {verifyCaCertificate} from "./verify";

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
        const days =  parseDays(command.days);
        const isCertificateGenerated = await generateCertificates(command.caCert, command.cert, command.key, days);
        if (isCertificateGenerated) {
            await uninstallCaCertificate();
            await installCaCertificate(command.caCert);
        }
    } catch (err) {
        logErrorMessage(err);
    }
}

export async function verify(command: commander.Command) {
    try {
        await verifyCaCertificate();
    } catch (err) {
        logErrorMessage(err);
    }
}

export async function uninstall(command: commander.Command) {
    try {
        await uninstallCaCertificate();
    } catch (err) {
        logErrorMessage(err);
    }
}
