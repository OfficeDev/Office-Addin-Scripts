// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as chalk from 'chalk';
import { parseNumber } from "office-addin-cli";
import { ManifestInfo, readManifestFile } from 'office-addin-manifest';
import { sendUsageDataException, sendUsageDataSuccessEvent } from './defaults';
import * as configure from './configure';
import { SSOService } from './server';
import { addSecretToCredentialStore, writeApplicationData } from './ssoDataSettings';

let usageDataInfo: Object = {};

export async function configureSSO(manifestPath: string) {
    // Check platform and return if not Windows or Mac
    if (process.platform !== "win32" && process.platform !== "darwin") {
        console.log(chalk.yellow(`${process.platform} is not supported. Only Windows and Mac are supported`));
        return;
    }

    const port: number = parseDevServerPort(process.env.npm_package_config_dev_server_port) || 3000;

    // Log start time for configuration process
    const ssoConfigStartTime = (new Date()).getTime();

    // Check to see if Azure CLI is installed.  If it isn't installed, then install it
    const cliInstalled = await configure.isAzureCliInstalled();

    if (!cliInstalled) {
        console.log(chalk.yellow("Azure CLI is not installed.  Installing now before proceeding - this could take a few minutes."));
        await configure.installAzureCli();
        if (process.platform === "win32") {
            console.log(chalk.green('Please close your command shell, reopen and run configure-sso again.  This is necessary to register the path to the Azure CLI'));
        }
        return;
    }

    console.log('Opening browser for authentication to Azure. Enter valid Azure credentials');
    const userJson: Object = await configure.logIntoAzure();
    if (Object.keys(userJson).length >= 1) {
        console.log('Login was successful!');
        const manifestInfo: ManifestInfo = await readManifestFile(manifestPath);

        // Register application in Azure
        console.log('Registering new application in Azure');
        const applicationJson: Object = await configure.createNewApplication(manifestInfo.displayName, port.toString(), userJson);

        if (applicationJson) {
            console.log('Application was successfully registered with Azure');
            // Set application IdentifierUri
            console.log('Setting identifierUri');
            await configure.setIdentifierUri(applicationJson, port.toString());

            // Set application sign-in audience
            console.log('Setting signin audience');
            await configure.setSignInAudience(applicationJson);

            // Grant admin consent for application if logged-in user is a tenant admin
            if (await configure.isUserTenantAdmin(userJson)) {
                console.log('Granting admin consent');
                await configure.grantAdminContent(applicationJson);
                // Check to set if SharePoint reply urls are set for tenant. If not, set them
                const setSharePointReplyUrls: boolean = await configure.setSharePointTenantReplyUrls();
                if (setSharePointReplyUrls) {
                    console.log('Set SharePoint reply urls for tenant');
                }
                // Check to set if Outlook reply url is set for tenant. If not, set them
                const setOutlookReplyUrl: boolean = await configure.setOutlookTenantReplyUrl();
                if (setOutlookReplyUrl) {
                    console.log('Set Outlook reply url for tenant');
                }
            }

            // Create an application secret and add to the credential store
            console.log('Setting application secret');
            const secret: string = await configure.setApplicationSecret(applicationJson);
            console.log(chalk.green(`App secret is ${secret}`));

            // Add secret to Credential Store (Windows) or Keychain(Mac)
            if (process.platform === "win32") {
                console.log(`Adding application secret for ${manifestInfo.displayName} to Windows Credential Store`);
            }
            else {
                console.log(`Adding application secret for ${manifestInfo.displayName} to Mac OS Keychain. You will need to provide an admin password to update the Keychain`);
            }
            addSecretToCredentialStore(manifestInfo.displayName, secret);
        } else {
            const errorMessage = 'Failed to register application';
            sendUsageDataException('createNewApplication', errorMessage);
            console.log(chalk.red(errorMessage));
            return;
        }
        // Write application data to project files (manifest.xml, .env, src/taskpane/fallbacktaskpane.ts)
        console.log(`Updating source files with application ID and port`);
        const projectUpdated = await writeApplicationData(applicationJson['appId'], port.toString(), manifestPath);
        if (!projectUpdated) {
            console.log(chalk.yellow(`Project was already previously updated. You will need to update the CLIENT_ID and PORT settings manually`));
        }

        // Log out of Azure
        console.log('Logging out of Azure now');
        await configure.logoutAzure();
        console.log(chalk.green(`Application with id ${applicationJson['appId']} successfully registered in Azure.  Go to https://ms.portal.azure.com/#home and search for 'App Registrations' to see your application`));

        // Log end time for configuration process and compute in seconds
        const ssoConfigEndTime = (new Date()).getTime();
        const ssoConfigDuration = (ssoConfigEndTime - ssoConfigStartTime) / 1000

        // Send usage data
        sendUsageDataSuccessEvent('configureSSO', {configDuration: ssoConfigDuration});
    }
    else {
        const errorMessage: string = 'Login to Azure did not succeed';
        sendUsageDataException('configureSSO', errorMessage);
        throw new Error(errorMessage);
    }

}

export async function startSSOService(manifestPath: string) {
    // Check platform and return if not Windows or Mac
    if (process.platform !== "win32" && process.platform !== "darwin") {
        console.log(chalk.yellow(`${process.platform} is not supported. Only Windows and Mac are supported`));
        return;
    }
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