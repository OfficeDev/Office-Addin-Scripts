// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as cors from "cors";
import * as devCerts from "office-addin-dev-certs";
import * as express from "express";
import * as path from "path";
import * as https from "https";

export const defaultPort: number = 4201;

export class TestServer {
    m_jsonData: any;
    m_port: number;
    m_testServerStarted: boolean;
    m_app: express.Express;
    m_resultsPromise: Promise<JSON>;
    m_server: https.Server;

    constructor(port: number) {
        this.m_app = express();
        this.m_jsonData = {};
        this.m_port = port;
        this.m_resultsPromise = undefined;
        this.m_testServerStarted = false;
    }

    public async startTestServer(mochaTest: boolean = false): Promise<any> {
        try {
            if (mochaTest) {
                process.env["NODE_TLS_REJECT_UNAUTHORIZED"] = "0";
            }

            // create express server instance
            const options = await devCerts.getHttpsServerOptions();
            this.m_app.use(cors());
            this.m_server = https.createServer(options, this.m_app);

            // listen for 'ping'
            const platformName = this.getPlatformName();
            this.m_app.get("/ping", function (req: any, res: any, next: any) {
                res.send(platformName);
            });

            // listen for posting of test results
            this.m_resultsPromise = new Promise<JSON>(async (resolveResults) => {
                this.m_app.post("/results", async (req: any, res: any) => {
                    res.send("200");
                    this.m_jsonData = JSON.parse(req.query.data);
                    resolveResults(this.m_jsonData);
                });
            });

            // start listening on specified port
            return await this.startListening();

        } catch (err) {
            throw new Error(`Unable to start test server.\n${err}`);
        }
    }

    private async startListening(): Promise<any> {
        return new Promise<any>(async (resolve, reject) => {
            try {
                // set server to listen on specified port
                this.m_server.listen(this.m_port, () => {
                    this.m_testServerStarted = true;
                    resolve(true);
                });

            } catch (err) {
                reject(new Error(`Unable to start test server.\n${err}`));
            }
        });
    }

    public async stopTestServer(): Promise<boolean> {
        return new Promise<boolean>(async (resolve, reject) => {
            if (this.m_testServerStarted) {
                try {
                    this.m_server.close();
                    this.m_testServerStarted = false;
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
        return this.m_resultsPromise;
    }

    public getTestServerState(): boolean {
        return this.m_testServerStarted;
    }

    public getTestServerPort(): number {
        return this.m_port;
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
}
