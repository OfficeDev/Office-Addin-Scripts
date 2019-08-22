// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as cors from "cors";
import * as express from "express";
import * as http from "http";
import * as https from "https";
import * as devCerts from "office-addin-dev-certs";
import * as defaults from "./defaults";

export class TestServer {
    private jsonData: any;
    private httpPort: number;
    private httpsPort: number;
    private testServerStarted: boolean;
    private app: express.Express;
    private resultsPromise: Promise<JSON>;
    private httpServer: http.Server;
    private httpsServer: https.Server;

    /**
     * Creates the test server object and initializes member variables
     * @param port Specifies port the test server will run on
     */
    constructor(httpsPort: number = defaults.httpsPort, httpPort: number = defaults.httpPort) {
        this.app = express();
        this.jsonData = {};
        this.httpsPort = httpsPort;
        this.httpPort = httpPort;
        this.resultsPromise = undefined;
        this.testServerStarted = false;
    }

    /**
     * Start the test server
     * @param mochaTest @deprecated
     */
    public async startTestServer(mochaTest = false): Promise<boolean> {
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

    /**
     * Stop the test server if it's started
     */
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
    /**
     * Returns the test results sent to the test server
     */
    public async getTestResults(): Promise<JSON> {
        return this.resultsPromise;
    }

    /**
     * Returns whether the server is started or stopped
     */
    public getTestServerState(): boolean {
        return this.testServerStarted;
    }

    /**
     * Returns the https port for the test server
     */
    public getTestHttpsServerPort(): number {
        return this.httpsPort;
    }

    /**
     * Returns the http port for the test server
     */
    public getTestHttpServerPort(): number {
        return this.httpPort;
    }

    /**
     * Returns the OS platform the test server is running on
     */
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
