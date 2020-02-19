/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in root of repo. -->
 *
 * This file defines Azure application registration.
 */
import * as childProcess from 'child_process';
import * as defaults from './defaults';
import * as fs from 'fs';
import { usageDataObject } from './defaults';
require('dotenv').config();

export async function createNewApplication(ssoAppName: string, port: string, userJson: Object): Promise<Object> {
    try {
        let azRestCommand = await fs.readFileSync(defaults.azRestAppCreateCommandPath, 'utf8');
        const reName = new RegExp('{SSO-AppName}', 'g');
        const rePort = new RegExp('{PORT}', 'g');
        azRestCommand = azRestCommand.replace(reName, ssoAppName).replace(rePort, port);
        const applicationJson: Object = await promiseExecuteCommand(azRestCommand, true /* returnJson */);
        return applicationJson;
    } catch (err) {
        const errorMessage: string = `Unable to register new application: \n${err}`
        throw new Error(errorMessage);
    }
}

async function applicationReady(applicationJson: Object): Promise<boolean> {
    try {
        const azRestCommand: string = `az ad app show --id ${applicationJson['appId']}`
        const appJson: Object = await promiseExecuteCommand(azRestCommand, true /* returnJson */, true /* expectError */);
        return appJson !== "";
    } catch (err) {
        const errorMessage: string = `Unable to get application info: \n${err}`
        throw new Error(errorMessage);
    }
}

export async function grantAdminContent(applicationJson: Object): Promise<void> {
    try {
        // Check to see if the application is available before granting admin consent
        let appReady: boolean = false;
        let counter: number = 0;
        while (appReady === false && counter <= 50) {
            appReady = await applicationReady(applicationJson);
            counter++;
        }

        if (counter > 50) {
            const warningMessage: string = 'Application does not appear to be ready to grant admin consent';
            return;
        }

        const azRestCommand: string = `az ad app permission admin-consent --id ${applicationJson['appId']}`;
        await promiseExecuteCommand(azRestCommand);
    } catch (err) {
        const errorMessage: string = `Unable to set grant admin consent: \n${err}`
        throw new Error(errorMessage);
    }
}

export async function isAzureCliInstalled(): Promise<boolean> {
    try {
        let cliInstalled: boolean = false;
        switch (process.platform) {
            case "win32":
                const appsInstalledWindowsCommand: string = `powershell -ExecutionPolicy Bypass -File "${defaults.getInstalledAppsPath}"`;
                const appsWindows: any = await promiseExecuteCommand(appsInstalledWindowsCommand);
                cliInstalled = appsWindows.filter(app => app.DisplayName === 'Microsoft Azure CLI').length > 0;
                // Send usage data
                usageDataObject.sendUsageDataSuccessEvent('isAzureCliInstalled', {cliInstalled: cliInstalled});
                return cliInstalled;
            case "darwin":
                const appsInstalledMacCommand = 'brew list';
                const appsMac: Object | string = await promiseExecuteCommand(appsInstalledMacCommand, false /* returnJson */);
                cliInstalled = appsMac.toString().includes('azure-cli');
                // Send usage data
                usageDataObject.sendUsageDataSuccessEvent('isAzureCliInstalled', {cliInstalled: cliInstalled});
                return cliInstalled;
            default:
                const errorMessage: string = `Platform not supported: ${process.platform}`;
                throw new Error(errorMessage);
        }        
    } catch (err) {
        const errorMessage: string = `Unable to determine if Azure CLI is installed: \n${err}`;
        throw new Error(errorMessage);
    }
}

export async function installAzureCli(): Promise<void> {
    try {
        switch (process.platform) {
            case "win32":
                const windowsCliInstallCommand = `powershell -ExecutionPolicy Bypass -File "${defaults.azCliInstallCommandPath}"`;
                await promiseExecuteCommand(windowsCliInstallCommand, false /* returnJson */);
                break;
            case "darwin": // macOS
                const macCliInstallCommand = 'brew update && brew install azure-cli';
                await promiseExecuteCommand(macCliInstallCommand, false /* returnJson */);
                break;
            default:
                const errorMessage: string = `Platform not supported: ${process.platform}`;
                throw new Error(errorMessage);
        }
    } catch (err) {
        const errorMessage: string = `Unable to install Azure CLI: \n${err}`;
        throw new Error(errorMessage);
    }
}

export async function isUserTenantAdmin(userInfo: Object): Promise<boolean> {
    try {
        let azRestCommand: string = fs.readFileSync(defaults.azRestGetTenantRolesPath, 'utf8');
        const tenantRoles: Object = await promiseExecuteCommand(azRestCommand);
        let tenantAdminId: string = '';
        tenantRoles['value'].forEach(item => {
            if (item.displayName === "Company Administrator") {
                tenantAdminId = item.id;
            }
        });

        azRestCommand = fs.readFileSync(defaults.azRestGetTenantAdminMembershipCommandPath, 'utf8');
        azRestCommand = azRestCommand.replace('<TENANT-ADMIN-ID>', tenantAdminId);
        const tenantAdmins: Object = await promiseExecuteCommand(azRestCommand);
        let isTenantAdmin: boolean = false;
        tenantAdmins['value'].forEach(item => {
            if (item.userPrincipalName === userInfo[0].user.name) {
                isTenantAdmin = true;
            }
        });

        // Send usage data
        usageDataObject.sendUsageDataSuccessEvent('isUserTenantAdmin', {isUserTenantAdmin: isTenantAdmin});

        return isTenantAdmin;
    } catch (err) {
        const errorMessage: string = `Unable to determine if user is tenant admin: \n${err}`;
        throw new Error(errorMessage);
    }
}

export async function logIntoAzure(): Promise<Object> {
    let userJson: Object = await promiseExecuteCommand('az login --allow-no-subscriptions', true /* returnJson */, true /* expectError */);
    if (Object.keys(userJson).length < 1) {
        // Try alternate login
        logoutAzure();
        userJson = await promiseExecuteCommand('az login');
    }
    return userJson
}

export async function logoutAzure(): Promise<Object> {
    return await promiseExecuteCommand('az logout');
}

async function promiseExecuteCommand(cmd: string, returnJson: boolean = true, expectError: boolean = false): Promise<Object | string> {
    return new Promise((resolve, reject) => {
        try {
            childProcess.exec(cmd, { maxBuffer: 1024 * 102400 }, async (err, stdout, stderr) => {
                if (err && !expectError) {
                    reject(stderr);
                }
                
                let results = stdout;
                if (results !== '' && returnJson) {
                    results = JSON.parse(results);
                }
                resolve(results);
            });
        } catch (err) {
            reject(err);
        }
    });
}

export async function setApplicationSecret(applicationJson: Object): Promise<string> {
    try {
        let azRestCommand: string = await fs.readFileSync(defaults.azRestAddSecretCommandPath, 'utf8');
        azRestCommand = azRestCommand.replace('<App_Object_ID>', applicationJson['id']);
        const secretJson: Object = await promiseExecuteCommand(azRestCommand);
        return secretJson['secretText'];
    } catch (err) {
        const errorMessage: string = `Unable to set application secret: \n${err}`;
        throw new Error(errorMessage);
    }
}

export async function setIdentifierUri(applicationJson: Object, port: string): Promise<void> {
    try {
        let azRestCommand: string = await fs.readFileSync(defaults.azRestSetIdentifierUriCommmandPath, 'utf8');
        azRestCommand = azRestCommand.replace('<App_Object_ID>', applicationJson['id']).replace('<App_Id>', applicationJson['appId']).replace('{PORT}', port.toString());
        await promiseExecuteCommand(azRestCommand);
    } catch (err) {
        const errorMessage: string = `Unable to set identifierUri: \n${err}`;
        throw new Error(errorMessage);
    }
}

export async function setSignInAudience(applicationJson: Object):Promise<void> {
    try {
        let azRestCommand: string = fs.readFileSync(defaults.azRestSetSigninAudienceCommandPath, 'utf8');
        azRestCommand = azRestCommand.replace('<App_Object_ID>', applicationJson['id']);
        await promiseExecuteCommand(azRestCommand);
    } catch (err) {
        const errorMessage: string = `Unable to set signInAudience: \n${err}`;
        throw new Error(errorMessage);
    }
}

export async function setSharePointTenantReplyUrls(): Promise<boolean> {
    try {
        let servicePrinicipaObjectlId = "";
        let setReplyUrls: boolean = true;
        const sharePointServiceId: string = "57fb890c-0dab-4253-a5e0-7188c88b2bb4";

        // Get tenant name and construct SharePoint SSO reply urls with tenant name
        let azRestCommand: string = fs.readFileSync(defaults.azRestGetOrganizationDetailsCommandPath, 'utf8');
        const tenantDetails: any = await promiseExecuteCommand(azRestCommand);
        const tenantName: string = tenantDetails.value[0].displayName
        const oneDriveReplyUrl: string = `https://${tenantName}-my.sharepoint.com/_forms/singlesignon.aspx`;
        const sharePointReplyUrl: string = `https://${tenantName}.sharepoint.com/_forms/singlesignon.aspx`;

        // Get service principals for tenant
        azRestCommand = 'az ad sp list --all';
        const servicePrincipals: any = await promiseExecuteCommand(azRestCommand);

        // Check if SharePoint redirects are set for SharePoint principal
        for (let item of servicePrincipals) {
            if (item.appId === sharePointServiceId) {
                servicePrinicipaObjectlId = item.objectId;
                if (item.replyUrls.length === 0) {
                    break;
                    // if there are reply urls set, then we need to see if the SharePoint SSO reply urls are already set
                } else {
                    for (let url of item.replyUrls) {
                        if (url === oneDriveReplyUrl || url === sharePointReplyUrl) {
                            setReplyUrls = false;
                            break;
                        }
                    }
                }

            }
        }

        if (setReplyUrls) {
            azRestCommand = fs.readFileSync(defaults.azRestAddTenantReplyUrlsCommandPath, 'utf8');
            const reName = new RegExp('<TENANT-NAME>', 'g');
            azRestCommand = azRestCommand.replace(reName, tenantName).replace('<SP-OBJECTID>', servicePrinicipaObjectlId);
            await promiseExecuteCommand(azRestCommand);
        }

        // Send usage data
        usageDataObject.sendUsageDataSuccessEvent('setTenantReplyUrls', {isUserTenantAdmin: setReplyUrls});
        return setReplyUrls;
    } catch (err) {
        const errorMessage: string = `Unable to set tenant reply urls. \n${err}`;
        throw new Error(errorMessage);
    }
}

export async function setOutlookTenantReplyUrl(): Promise<boolean> {
    try {
        let servicePrinicipaObjectlId = "";
        let setReplyUrls: boolean = true;
        const outlookReplyUrl: string = "https://outlook.office.com/owa/extSSO.aspx"
        const outlookServiceId = "bc59ab01-8403-45c6-8796-ac3ef710b3e3";

        // Get service principals for tenant
        let azRestCommand: string = 'az ad sp list --all';
        const servicePrincipals: any = await promiseExecuteCommand(azRestCommand);

        // Check if Outlook redirects are set for Outlook principal
        for (let item of servicePrincipals) {
            if (item.appId === outlookServiceId) {
                servicePrinicipaObjectlId = item.objectId;
                if (item.replyUrls.length === 0) {
                    break;
                    // if there are reply urls set, then we need to see if the Outlook SSO reply urls are already set
                } else {
                    for (let url of item.replyUrls) {
                        if (url === outlookReplyUrl) {
                            setReplyUrls = false;
                            break;
                        }
                    }
                }

            }
        }

        if (setReplyUrls) {
            azRestCommand = fs.readFileSync(defaults.azRestAddTenantOutlookReplyUrlsCommandPath, 'utf8');
            azRestCommand = azRestCommand.replace('<SP-OBJECTID>', servicePrinicipaObjectlId);
            await promiseExecuteCommand(azRestCommand);
        }

        // Send usage data
        usageDataObject.sendUsageDataSuccessEvent('setOutlookTenantReplyUrls', {tenantReplyUrlsSet: setReplyUrls});
        return setReplyUrls;
    } catch (err) {
        const errorMessage: string = `Unable to set tenant reply urls. \n${err}`;
        usageDataObject.sendUsageDataException('setOutlookTenantReplyUrls', errorMessage);
        throw new Error(errorMessage);
    }
}
