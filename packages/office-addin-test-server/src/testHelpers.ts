import * as childProcess from "child_process";
import * as cps from "current-processes";
import * as testServer from "../src/testServer";
const XMLHttpRequest = require("xmlhttprequest").XMLHttpRequest;
let devServerStarted: boolean = false;
let port: number = 8080;
var subProcess: any;

export async function setupTestEnvironment(app: string, port: number): Promise<boolean> {
    return new Promise<boolean>(async function (resolve, reject) {

        if (process.platform !== 'win32' && process.platform !== 'darwin') {
            reject();
        }

        devServerStarted = await _startDevServer();
        let sideLoadSuceeded: boolean = false
        try {
            console.log(`Sideload application: ${app}`);
            const cmdLine = `npm run sideload:${app}`;
            sideLoadSuceeded = await _executeCommandLine(cmdLine);
        } catch (err) {
            console.log(`Unable to sideload ${app}. ${err}`);
        }
        resolve(devServerStarted && sideLoadSuceeded);
    });
}

async function _startDevServer(): Promise<boolean> {
    devServerStarted = false;
    const cmdLine = "npm run dev-server";
    subProcess = childProcess.spawn(cmdLine, [], {
        detached: true,
        shell: true,
        stdio: "ignore"
    });
    subProcess.on("error", (err) => {
    console.log(`Unable to run command: ${cmdLine}.\n${err}`);
    });
    return subProcess.pid != undefined;
}

async function _executeCommandLine(cmdLine: string): Promise<boolean> {
    return new Promise<boolean>((resolve, reject) => {
        childProcess.exec(cmdLine, (error) => {
            if (error) {
                reject(false);
            } else {
                resolve(true);
            }
        });
    });
}

export async function teardownTestEnvironment(processName: string): Promise<void> {
    const operatingSystem: string = process.platform;
    try {
        if (operatingSystem == 'win32') {
            const cmdLine = `tskill ${processName}`;
            await _executeCommandLine(cmdLine);
        } else {
            const pid = await _getProcessId(processName);
            if (pid != undefined) {
                process.kill(pid);
            }
        }
    } catch (err) {
        console.log(`Unable to kill Excel process. ${err}`);
    }

    // if the dev-server was started, kill the spawned process
    if (devServerStarted) {
        if (operatingSystem == 'win32') {
            childProcess.spawn("taskkill", ["/pid", subProcess.pid, '/f', '/t']);
        } else {
            subProcess.kill();
        }
    }
}

async function _getProcessId(processName: string): Promise<number> {
    return new Promise<number>(async function (resolve) {
        cps.get(function (err: Error, processes: any) {
            try {
                const p = processes.filter(function (p: any) {
                    return (p.name.indexOf(processName) > 0);
                });
                resolve(p.length > 0 ? p[0].pid : undefined);
            }
            catch (err) {
                console.log("Unable to get list of processes");
            }
        });
    });
}

export async function pingTestServer(portNumber: number | undefined): Promise<Object> {
    return new Promise<Object>(async (resolve, reject) => {
        if (portNumber !== undefined) {
            port = portNumber;
        }

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

export async function sendTestResults(data: Object, portNumber: number | undefined): Promise<boolean> {
    return new Promise<boolean>(async (resolve, reject) => {
        if (portNumber !== undefined) {
            port = portNumber;
        }

        const json = JSON.stringify(data);
        const xhr = new XMLHttpRequest();
        const url: string = `https://localhost:${port}/results/`;
        const dataUrl: string = url + "?data=" + encodeURIComponent(json);

        xhr.onreadystatechange = () => {
            if (xhr.readyState === 4 && xhr.status === 200 && xhr.responseText === "200") {
                resolve(true);
            }
            else if (xhr.readyState === 4 && xhr.status === 0 && xhr.responseText.indexOf("ECONNREFUSED") > 0) {
                reject(false);
            }
        };
        xhr.open("POST", dataUrl, true);
        xhr.send();
    });
}