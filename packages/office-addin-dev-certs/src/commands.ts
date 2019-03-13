import * as commander from "commander";
import {generateCertificates} from "./generate";
import {installCaCertificate} from "./install";
import {uninstallCaCertificate} from "./uninstall";
import {verifyCaCertificate} from "./verify";

export async function generate(command: commander.Command) {
    try {
        await generateCertificates(command.caCert, command.cert, command.key, parseInt(command.days, 10), command.install);
    } catch (err) {
        console.error(`Unable to generate self-signed dev certificates.\n${err}`);
    }
}

export async function install(command: commander.Command) {
    try {
        await installCaCertificate(command.caCert);
    } catch (err) {
        console.error(`Unable to install the CA certificate.\n${err}`);
    }
}

export async function verify(command: commander.Command) {
    try {
        await verifyCaCertificate();
    } catch (err) {
        console.error(`Unable to verify CA certificates.\n${err}`);
    }
}

export async function uninstall(command: commander.Command) {
    try {
        await uninstallCaCertificate();
    } catch (err) {
        console.error(`Unable to uninstall CA certificates.\n${err}`);
    }
}
