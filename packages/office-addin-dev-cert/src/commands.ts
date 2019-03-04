import * as commander from "commander";
import {generateCertificates} from "./generate"
import {installCertificates} from "./install"
import {uninstallCertificates} from "./uninstall"
import {cleanCertificates} from "./clean"
import {verifyCertificates} from "./verify"

function generatePlatformDependentPath(path: string | undefined): string{
    const certFolder  = "certs"; //read from manifest file?
    const folderPrefix = (process.platform == "win32")? ".\\" : "./";
    if (path == undefined) { 
        path = folderPrefix + certFolder;
    }else{
        //add to manifest to file
    }
    return path;
}

export async function generate(manifestPath: string, command: commander.Command) {
    try {
        await generateCertificates(generatePlatformDependentPath(command.path));
    } catch (err) {
        console.error(`Unable to generate self-signed dev certificates.\n${err}`);
    }
}

export async function install(manifestPath: string, command: commander.Command) {
    try {
        await installCertificates(generatePlatformDependentPath(command.path));
    } catch (err) {
        console.error(`Unable to install dev certificates.\n${err}`);
    }
}

export async function verify(manifestPath: string, command: commander.Command) {
    try {
        await verifyCertificates(generatePlatformDependentPath(command.path));
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

export async function clean(manifestPath: string, command: commander.Command) {
    try {
        await cleanCertificates(generatePlatformDependentPath(command.path));
    } catch (err) {
        console.error(`Unable to install self-signed dev certificates.\n${err}`);
    }
}

