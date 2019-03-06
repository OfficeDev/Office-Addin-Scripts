import * as cors from "cors";
import * as express from "express";
import * as fs from "fs";
const path = require("path");

export class TestServer {
    m_jsonData: any;
    m_port: number;
    m_testServerStarted;
    m_app: any;
    m_resultsPromise: Promise<JSON>;
    m_server: any;

    constructor(port: number) {
        this.m_app = express();
        this.m_jsonData = {};
        this.m_port = port;
        this.m_resultsPromise = undefined;
        this.m_testServerStarted = false;
    }

    public async startTestServer(mochaTest: boolean = false): Promise<any>{
        return new Promise<boolean>(async (resolve, reject) => {
            if (mochaTest) {
                process.env["NODE_TLS_REJECT_UNAUTHORIZED"] = "0";
            }

            const key = fs.readFileSync(path.resolve(__dirname, '../certs/server.key'));
            const cert = fs.readFileSync(path.resolve(__dirname, '../certs/server.crt'));
            const options = { key: key, cert: cert };

            // listen for 'ping'
            this.m_app.use(cors());
            this.m_app.get("/ping", function (req: any, res: any, next: any) {
                res.send(process.platform === "win32" ? "Win32" : "Mac");
            });

            // listen for posting of test results
            this.m_resultsPromise = new Promise<JSON>(async (resolveResults) => {
                this.m_app.post("/results", async function (req: any, res: any) {
                    res.send("200");
                    this.m_jsonData = JSON.parse(req.query.data);
                    resolveResults(this.m_jsonData);
                }.bind(this));

            });
    
            const https = require("https");
            this.m_server = https.createServer(options,  this.m_app);
    
            // start listening on specified port
            try {
                this.m_server.listen(this.m_port, function () {
                    this.m_testServerStarted = true;
                    resolve(true);
                }.bind(this));
            } catch (err) {
                reject(new Error(`Unable to start test server. \n${err}`));
            }
        });
    }
    
    public async stopTestServer(): Promise <boolean> {
        return new Promise<boolean>(async (resolve, reject) => {
            if (this.m_testServerStarted) {
                try {
                    this.m_server.close();
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
   
    public async getTestResults(): Promise<JSON> {
        return this.m_resultsPromise;
    }
    
    public getTestServerState(): boolean {
        return this.m_testServerStarted;
    }
    
    public getTestServerPort(): number {
        return this.m_port;
    }
}
