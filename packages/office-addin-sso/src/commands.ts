// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import { configureSSOApplication } from './configure';
import { SSOService } from './server';
import { parseNumber } from "office-addin-cli";

export async function configureSSO(manifestPath: string) {
    const port: number = parseDevServerPort(process.env.npm_package_config_dev_server_port) || 3000;
    await configureSSOApplication(manifestPath, port.toString());
}

export async function startSSOService(manifestPath: string) {
    const sso = new SSOService(manifestPath);
    sso.startSsoService();
}

function parseDevServerPort(optionValue: any): number | undefined {
    const devServerPort = parseNumber(optionValue, "--dev-server-port should specify a number.");

    if (devServerPort !== undefined) {
        if (!Number.isInteger(devServerPort)) {
            throw new Error("--dev-server-port should be an integer.");
        }
        if ((devServerPort < 0) || (devServerPort > 65535)) {
            throw new Error("--dev-server-port should be between 0 and 65535.");
        }
    }

    return devServerPort;
}