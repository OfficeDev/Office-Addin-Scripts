import * as cors from "cors";
import * as express from "express";
import * as fs from "fs";

export default class TestServer {
    m_jsonData: any;
    m_port: number;
    m_testServerStarted;
    m_app: any;
    m_server: any;

    constructor(port: number) {
        this.m_jsonData = {};
        this.m_port = port;
        this.m_testServerStarted = false;
        this.m_app = express();
    }

    public async startTestServer(): Promise <boolean> {
        return new Promise<boolean>(async (resolve, reject) => {   
      
            process.env["NODE_TLS_REJECT_UNAUTHORIZED"] = "0";
            const key = fs.readFileSync('certs/server.key');
            const cert = fs.readFileSync('certs/server.crt');
            const options = { key: key, cert: cert };
    
            this.m_app.use(cors());
            this.m_app.get("/ping", function (req: any, res: any, next: any) {
                res.send(process.platform === "win32" ? "Win32" : "Mac");
                resolve(true);
            });
    
            this.m_app.post("/results", function (req: any, res: any) {
                res.send("200");
                this.m_jsonData = JSON.parse(req.query.data);
                resolve(true);
            }.bind(this));
    
            const https = require("https");
            this.m_server = https.createServer(options,  this.m_app);
    
            // listen for new web clients:
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
    
    public async stopTestServer(): Promise < boolean > {
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
        return this.m_jsonData;
    }
    
    public getTestServerState(): boolean {
        return this.m_testServerStarted;
    }
    
    public getTestServerPort(): number {
        return this.m_port;
    }
}
