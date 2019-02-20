import * as cors from "cors";
import * as express from "express";
import * as fs from "fs";
const app: any = express();
let server: any;
const XMLHttpRequest = require("xmlhttprequest").XMLHttpRequest;

export default class officeAddinTestServer {
    m_port: number = 0;
    m_testServerStarted: boolean = false;
    m_jsonData: JSON = undefined;

    constructor(port: number) {
        this.m_port = port;
    }

    async startTestServer(): Promise<boolean> {
        return new Promise<boolean>(async (resolve, reject) => {
            process.env["NODE_TLS_REJECT_UNAUTHORIZED"] = "0";
            const key = fs.readFileSync('certs/server.key');
            const cert = fs.readFileSync('certs/server.crt');
            const options = { key: key, cert: cert };

            app.use(cors());
            app.get("/ping", function (req: any, res: any, next: any) {
                res.send(process.platform === "win32" ? "Win32" : "Mac");
                resolve(true);
            });

            app.post("/results", function (req: any, res: any) {
                res.send("200");
                this.m_jsonData = JSON.parse(req.query.data);
                resolve(true);
            }.bind(this));

            const https = require("https");
            server = https.createServer(options, app);

            // listen for new web clients:
            try {
                server.listen(this.m_port, function () {
                    this.m_testServerStarted = true;
                    resolve(true);
                }.bind(this));
            } catch (err) {
                reject(new Error(`Unable to start test server. \n${err}`));
            }
        });
    }

    async stopTestServer(): Promise<boolean> {
        return new Promise<boolean>(async (resolve, reject) => {
            if (this.m_testServerStarted) {
                try {
                    server.close();
                    this.m_testServerStarted = false;
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

    getTestResults(): JSON {
        return this.m_jsonData;
    }

    getTestServerState(): boolean {
        return this.m_testServerStarted;
    }

    getTestServerPort(): number {
        return this.m_port;
    }
}

export async function pingTestServer(port: number): Promise < string > {
    return new Promise<string>(async (resolve, reject) => {
        const xhr = new XMLHttpRequest();
        const pingUrl: string = `https://localhost:${port}/ping`;
        xhr.onreadystatechange = () => {
            if (xhr.readyState === 4 && xhr.status === 200) {
                resolve(xhr.responseText);
            }
            else if (xhr.readyState === 4 && xhr.status === 0 && xhr.responseText.indexOf("ECONNREFUSED") > 0) {
                reject(xhr.responseText);
            }
        };
        xhr.open("GET", pingUrl, true);
        xhr.send();
    });
}

export async function sendTestResults(data: any, port: number): Promise<boolean> {
    return new Promise<boolean>(async (resolve, reject) => {
        const json = JSON.stringify(data);
        const xhr = new XMLHttpRequest();
        const url: string = `https://localhost:${port}/results/`;
        const dataUrl: string = url + "?data=" + encodeURIComponent(json);

        xhr.open("POST", dataUrl, true);
        xhr.send();
        xhr.onreadystatechange = () => {
            if (xhr.readyState === 4 && xhr.status === 200 && xhr.responseText === "200") {
                resolve(true);
            }
            else if (xhr.readyState === 4 && xhr.status === 0 && xhr.responseText.indexOf("ECONNREFUSED") > 0) {
                reject(false);
            }
        };
    });
}