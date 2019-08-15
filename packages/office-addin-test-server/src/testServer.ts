// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as cors from "cors";
import * as express from "express";
import * as http from "http";
import * as https from "https";
import * as devCerts from "office-addin-dev-certs";

export const defaultHttpPort: number = 4200;
export const defaultHttpsPort: number = 4201;

export class TestServer {
    private jsonData: any;
    private httpPort: number;
    private httpsPort: number;
    private testServerStarted: boolean;
    private app: express.Express;
    private resultsPromise: Promise<JSON>;
    private httpServer: http.Server;
    private httpsServer: https.Server;

    constructor(port: number) {
        this.app = express();
        this.jsonData = {};
        this.httpsPort = port;
        this.httpPort = defaultHttpPort;
        this.resultsPromise = undefined;
        this.testServerStarted = false;
    }

    public async startTestServer(mochaTest = false /* deprecated */): Promise<boolean> {
        try {
            // create express server instance
            const options = await devCerts.getHttpsServerOptions();
            this.app.use(cors());
            this.app.use(express.json());
            this.httpServer = http.createServer(this.app);
            this.httpsServer = https.createServer(options, this.app);

            // listen for 'ping'
            const platformName = this.getPlatformName();
            this.app.get("/ping", function(req: any, res: any, next: any) {
                res.send(platformName);
            });

            // listen for posting of test results
            this.resultsPromise = new Promise<JSON>(async (resolveResults) => {
                this.app.post("/results", async (req: any, res: any) => {
                    res.send("200");
                    this.jsonData = req.body;
                    resolveResults(this.jsonData);
                });
            });

            // start listening on specified port
            return await this.startListening();

        } catch (err) {
            throw new Error(`Unable to start test server.\n${err}`);
        }
    }

    public async stopTestServer(): Promise<boolean> {
        return new Promise<boolean>(async (resolve, reject) => {
            if (this.testServerStarted) {
                try {
                    this.httpServer.close();
                    this.httpsServer.close();
                    this.testServerStarted = false;
                    resolve(true);
                } catch (err) {
                    reject(new Error(`Unable to stop test server.\n${err}`));
                }
            } else {
                // test server not started
                resolve(false);
            }
        });
    }

    public async getTestResults(): Promise<JSON> {
        return this.resultsPromise;
    }

    public getTestServerState(): boolean {
        return this.testServerStarted;
    }

    public getTestHttpsServerPort(): number {
        return this.httpsPort;
    }

    public getTestHttpServerPort(): number {
        return this.httpPort;
    }

    public getPlatformName(): string {
        switch (process.platform) {
            case "win32":
                return "Windows";
            case "darwin":
                return "macOS";
            default:
                return process.platform;
        }
    }

    private async startListening(): Promise<boolean> {
        return new Promise<boolean>(async (resolve, reject) => {
            try {
                // set server to listen on specified ports
                this.httpsServer.listen(this.httpsPort, () => {
                    this.httpServer.listen(this.httpPort, () => {
                        this.testServerStarted = true;
                        resolve(true);
                    });
                });

            } catch (err) {
                reject(new Error(`Unable to start test server.\n${err}`));
            }
        });
    }
}
