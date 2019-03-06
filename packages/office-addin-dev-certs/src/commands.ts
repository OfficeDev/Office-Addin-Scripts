import * as commander from "commander";
import {generateCertificates} from "./generate";
import {installCertificates} from "./install";
import {uninstallCertificates} from "./uninstall";
import {verifyCertificates} from "./verify";

export async function generate(command: commander.Command) {
    try {
        await generateCertificates(command.caCert || "ca.crt", command.cert || "localhost.crt", command.key || "localhost.key", command.overwrite, command.install);
    } catch (err) {
        console.error(`Unable to generate self-signed dev certificates.\n${err}`);
    }
}

export async function install(command: commander.Command) {
    try {
        await installCertificates(command.caCert);
    } catch (err) {
        console.error(`Unable to install dev certificates.\n${err}`);
    }
}

export async function verify(command: commander.Command) {
    try {
        await verifyCertificates();
    } catch (err) {
        console.error(`Unable to verify dev certificates.\n${err}`);
    }
}

export async function uninstall(command: commander.Command) {
    try {
        await uninstallCertificates();
    } catch (err) {
        console.error(`Unable to uninstall dev certificates.\n${err}`);
    }
}
