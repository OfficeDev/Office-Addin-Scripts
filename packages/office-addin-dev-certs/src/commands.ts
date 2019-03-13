import * as commander from "commander";
import * as generate from "./generate";
import * as install from "./install";
import * as uninstall from "./uninstall";
import * as verify from "./verify";

export async function generateCertificates(command: commander.Command) {
    try {
        await generate.generateCertificates(command.caCert, command.cert, command.key, command.install);
    } catch (err) {
        console.error(`Unable to generate self-signed dev certificates.\n${err}`);
    }
}

export async function installCaCertificate(command: commander.Command) {
    try {
        await install.installCaCertificate(command.caCert);
    } catch (err) {
        console.error(`Unable to install the CA certificate.\n${err}`);
    }
}

export async function verifyCaCertificate(command: commander.Command) {
    try {
        await verify.verifyCaCertificate();
    } catch (err) {
        console.error(`Unable to verify CA certificates.\n${err}`);
    }
}

export async function uninstallCaCertificate(command: commander.Command) {
    try {
        await uninstall.uninstallCaCertificate();
    } catch (err) {
        console.error(`Unable to uninstall CA certificates.\n${err}`);
    }
}
