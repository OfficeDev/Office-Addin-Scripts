import * as commander from "commander";
import {logErrorMessage, parseNumber} from "office-addin-cli";
import * as defaults from "./defaults";
import {ensureCertificatesAreInstalled} from "./install";
import * as lockFile from "./lockfile";
import {deleteCertificateFiles, uninstallCaCertificate} from "./uninstall";
import {verifyCertificates} from "./verify";

// To handle crtl-c
process.on("SIGINT", function() {
    process.exit();
});

// The process on exit, make sure to clean the lock file.
process.on("exit", function() {
    const lock = new lockFile.LockFile();
    lock.releaseLock();
});

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

        await ensureCertificatesAreInstalled(days, command.machine);
    } catch (err) {
        logErrorMessage(err);
    }
}

export async function verify(command: commander.Command) {
    try {
        if (await verifyCertificates()) {
            console.log(`You have trusted access to https://localhost.\nCertificate: ${defaults.localhostCertificatePath}\nKey: ${defaults.localhostKeyPath}`);
        } else {
            console.log(`You need to install certificates for trusted access to https://localhost.`);
        }
    } catch (err) {
        logErrorMessage(err);
    }
}

export async function uninstall(command: commander.Command) {
    try {
        await uninstallCaCertificate(command.machine);
        await deleteCertificateFiles(defaults.certificateDirectory);
    } catch (err) {
        logErrorMessage(err);
    }
}
