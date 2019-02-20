import * as cors from "cors";
import * as express from "express";
import * as fs from "fs";
let m_port: number = 0;
let m_testServerStarted: boolean = false;
let server: any;
const app: any = express();

export default class officeAddinTestInfrastructure {
    constructor(port: number) {
        m_port = port;
    }

    async startTestServer(): Promise<boolean> {
        return new Promise<boolean>(async function (resolve, reject) {
        process.env["NODE_TLS_REJECT_UNAUTHORIZED"] = "0";
        const key = fs.readFileSync('certs/server.key');
        const cert = fs.readFileSync('certs/server.crt');
        const options = { key: key, cert: cert };

        app.use(cors());
        app.get("/ping", function (req: any, res: any, next: any) {
            res.send(process.platform === "win32" ? "Win32" : "Mac");
            resolve(true);
        });

        const https = require("https");
        server = https.createServer(options, app);

        // listen for new web clients:
        try {
            server.listen(m_port, function () {
                m_testServerStarted = true;
                resolve(true);
            });
        } catch (err) {
            reject(new Error(`Unable to start test server. \n${err}`));
        }
    });
    }

    async stopTestServer(): Promise<boolean> {
        return new Promise<boolean>(async function (resolve, reject) {
            if (m_testServerStarted) {
                try {
                    server.close();
                    m_testServerStarted = false;
                    resolve(true);
                } catch (err) {
                    reject(new Error(`Unable to stop test server. \n${err}`));
                }
            } else {
                // test server not started
                resolve(false);
            }
        });
    }

    async getTestResults(): Promise<any> {
        return new Promise<any>(async function (resolve, reject) {
            try {
                app.get("/results", function (req: any, res: any) {
                    res.send("200");
                    const jsonData: JSON = JSON.parse(req.query.data);
                    resolve(jsonData);
                });
            } catch (err) {
                reject(new Error(`Unable to get test results. \n${err}`));
            }
        });
    }

    getTestServerState(): boolean {
        return m_testServerStarted;
    }

    getTestServerPort(): number {
        return m_port;
    }
}