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