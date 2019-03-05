import * as commander from "commander";
import {cleanCertificates} from "./clean";
import {generateCertificates} from "./generate";
import {installCertificates} from "./install";
import {uninstallCertificates} from "./uninstall";
import {verifyCertificates} from "./verify";

export async function generate(command: commander.Command) {
    try {
        await generateCertificates(command.path);
    } catch (err) {
        console.error(`Unable to generate self-signed dev certificates.\n${err}`);
    }
}

export async function install(command: commander.Command) {
    try {
        await installCertificates(command.path);
    } catch (err) {
        console.error(`Unable to install dev certificates.\n${err}`);
    }
}

export async function verify(command: commander.Command) {
    try {
        await verifyCertificates(command.path);
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

export async function clean(command: commander.Command) {
    try {
        await cleanCertificates(command.path);
    } catch (err) {
        console.error(`Unable to install self-signed dev certificates.\n${err}`);
    }
}

