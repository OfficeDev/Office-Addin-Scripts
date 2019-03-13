import * as commander from "commander";
import {generateCertificates} from "./generate";
import {installCaCertificate} from "./install";
import {uninstallCaCertificate} from "./uninstall";
import {verifyCaCertificate} from "./verify";

function parseNumericCommandOption(optionValue: any, errorMessage: string = "The value should be a number."): number | undefined {
    switch (typeof(optionValue)) {
        case "number": {
            return optionValue;
        }
        case "string": {
            let result;

            try {
                result = parseInt(optionValue, 10);
            } catch (err) {
                throw new Error(errorMessage);
            }

            if (Number.isNaN(result)) {
                throw new Error(errorMessage);
            }

            return result;
        }
        case "undefined": {
            return undefined;
        }
        default: {
            throw new Error(errorMessage);
        }
    }
}

function parseDays(optionValue: any): number | undefined {
    const daysUntilCertificateExpires = parseNumericCommandOption(optionValue, "Days should specify a number.");

    if (daysUntilCertificateExpires !== undefined) {
        if (daysUntilCertificateExpires <= 0) {
            throw new Error("Days should be greater than zero.");
        }
    }

    return daysUntilCertificateExpires;
}

export async function generate(command: commander.Command) {
    try {
        const daysUntilCertificateExpires =  parseDays(command.days);
        await generateCertificates(command.caCert, command.cert, command.key, daysUntilCertificateExpires, command.install);
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
