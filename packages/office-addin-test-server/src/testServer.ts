import * as cors from "cors";
import * as express from "express";
import * as fs from "fs";
const XMLHttpRequest = require("xmlhttprequest").XMLHttpRequest;
const app: any = express();
let port: number = 8080;
let server: any;
let testServerStarted: boolean = false
let jsonData: JSON;

export async function startTestServer(portNumber: number | undefined): Promise <boolean> {
    return new Promise<boolean>(async (resolve, reject) => {

        if (portNumber !== undefined) {
            port = portNumber;
        }
    
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
            jsonData = JSON.parse(req.query.data);
            resolve(true);
        });

        const https = require("https");
        server = https.createServer(options, app);

        // listen for new web clients:
        try {
            server.listen(port, function () {
                testServerStarted = true;
                resolve(true);
            });
        } catch (err) {
            reject(new Error(`Unable to start test server. \n${err}`));
        }
    });
}

export async function stopTestServer(): Promise < boolean > {
    return new Promise<boolean>(async (resolve, reject) => {
        if (testServerStarted) {
            try {
                server.close();
                testServerStarted = false;
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

export function getTestResults(): JSON {
    return jsonData;
}

export function getTestServerState(): boolean {
    return testServerStarted;
}

export function getTestServerPort(): number {
    return port;
}

export async function pingTestServer(): Promise<Object> {
    return new Promise<Object>(async (resolve, reject) => {

        const serverResponse: any = {};
        const serverStatus: string = "status";
        const platform: string = "platform";
        const xhr = new XMLHttpRequest();
        const pingUrl: string = `https://localhost:${port}/ping`;
        xhr.onreadystatechange = () => {
            if (xhr.readyState === 4 && xhr.status === 200) {
                serverResponse[serverStatus] = xhr.status;
                serverResponse[platform] = xhr.responseText;
                resolve(serverResponse);
            }
            else if (xhr.readyState === 4 && xhr.status === 0 && xhr.responseText.indexOf("ECONNREFUSED") > 0) {
                reject(xhr.responseText);
            }
        };
        xhr.open("GET", pingUrl, true);
        xhr.send();
    });
}

export async function sendTestResults(data: Object): Promise<boolean> {
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