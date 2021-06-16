// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.
/*
    This file provides server configuration, startup and stop
*/
import * as https from "https";
import * as devCerts from "office-addin-dev-certs";
import * as manifest from "office-addin-manifest";
import { App } from "./app";
import { getSecretFromCredentialStore } from "./ssoDataSettings";
import { usageDataObject } from "./defaults";

/* global require, process, console */

require("dotenv").config();

export class SSOService {
  private app: App;
  manifestPath: string;
  private port: number | string;
  private server: https.Server;
  private ssoServiceStarted: boolean;
  constructor(manifestPath: string) {
    this.port = process.env.PORT || "3000";
    this.app = new App(this.port);
    this.app.initialize();
    this.manifestPath = manifestPath;
    this.ssoServiceStarted = false;
  }

  public async startSsoService(isTest: boolean = false): Promise<boolean> {
    try {
      if (isTest) {
        process.env.NODE_TLS_REJECT_UNAUTHORIZED = "0";
      }
      await this.getSecret(isTest);
      await this.startServer(this.app.appInstance, this.port);
      this.ssoServiceStarted = true;
      usageDataObject.reportSuccess("startSsoService()");
      return Promise.resolve(true);
    } catch (err) {
      console.error(`Failed to start SSO server. ${err}`);
      usageDataObject.reportException("startSsoService()", err);
      return Promise.reject(false);
    }
  }

  private async getSecret(isTest: boolean = false): Promise<void> {
    const manifestInfo = await manifest.readManifestFile(this.manifestPath);
    const appSecret = getSecretFromCredentialStore(manifestInfo.displayName, isTest);
    if (appSecret === "") {
      const errorMessage: string = `Credential store does not have secret for '${manifestInfo.displayName}'.`;
      throw new Error(errorMessage);
    }
    process.env.secret = appSecret;
  }

  public getTestServerState(): boolean {
    return this.ssoServiceStarted;
  }

  public async startServer(app, port): Promise<boolean> {
    try {
      const options = await devCerts.getHttpsServerOptions();
      this.server = https.createServer(options, app).listen(port, () => console.log(`Server running on ${port}`));
      return Promise.resolve(true);
    } catch (err) {
      const errorMessage: string = `Unable to start test server on port ${port}: \n${err}`;
      return Promise.reject(errorMessage);
    }
  }

  public async stopServer(): Promise<boolean> {
    if (this.ssoServiceStarted) {
      try {
        this.server.close();
        this.ssoServiceStarted = false;
        return Promise.resolve(true);
      } catch (err) {
        const errorMessage: string = `Unable to stop test server: \n${err}`;
        return Promise.reject(new Error(errorMessage));
      }
    } else {
      // test server not started
      return Promise.resolve(false);
    }
  }
}
