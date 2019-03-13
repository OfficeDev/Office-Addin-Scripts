
import * as childProcess from "child_process";
import * as cps from "current-processes";

let devServerStarted: boolean = false;
let port: number = 8080;
var subProcess: any;
const XMLHttpRequest = require("xmlhttprequest").XMLHttpRequest;

export async function pingTestServer(portNumber: number | undefined): Promise<Object> {
    return new Promise<Object>(async (resolve, reject) => {
        if (portNumber !== undefined) {
            port = portNumber;
        }

        const serverResponse: any = {};
        const xhr = new XMLHttpRequest();
        const pingUrl: string = `https://localhost:${port}/ping`;
        xhr.onreadystatechange = () => {
            if (xhr.readyState === 4 && xhr.status === 200) {
                serverResponse["status"] = xhr.status;
                serverResponse["platform"] = xhr.responseText;
                resolve(serverResponse);
            }
            else if (xhr.readyState === 4 && xhr.status === 0 && xhr.statusText.message === "XHR error") {
                reject(serverResponse);
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
            else if (xhr.readyState === 4 && xhr.status === 0 && xhr.statusTest === "XHR error") {
                resolve(false);
            }
        };
        xhr.open("POST", dataUrl, true);
        xhr.send();
    });
}


export async function sideloadDesktopApp(application: string): Promise<boolean> {
    return new Promise<boolean>(async function (resolve, reject) {

        if (process.platform !== 'win32' && process.platform !== 'darwin') {
            reject(false);
        }

        try {
            const cmdLine = `npm run sideload:${application}`;
            const sideloadSucceeded = await _executeCommandLine(cmdLine);
            resolve(sideloadSucceeded);
        } catch (err) {
            reject(err);
        }
    });
}

export async function startDevServer(): Promise<boolean> {
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

    devServerStarted = subProcess.pid != undefined;

    return devServerStarted;
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

// Office-JS close workbook API is coming soon.  Once it's available we can stript out this code to kill the Excel process
export async function teardownTestEnvironment(application: string): Promise<void> {

    await closeDesktopApplication(application);

    // if the dev-server was started, kill the spawned process
    if (devServerStarted) {
        if (process.platform == "win32") {
            childProcess.spawn("taskkill", ["/pid", subProcess.pid, '/f', '/t']);
        } else {
            subProcess.kill();
        }
    }
}

async function closeDesktopApplication(application: string): Promise <void> {
    let processName: string = "";
    switch(application.toLowerCase()) {
        case "excel":
            processName = "Excel";
            break;
        case "powerpoint":
            processName = "Powerpnt";
            break;
        case "onenote":
            processName = "Onenote";
            break;
        case "outlook":
            processName = "Outlook";
            break;
        case "project":
            processName = "Project";
            break;
        case "word":
            processName = "Winword";
            break;
        default:
            throw new Error(`${application} is not a valid Office desktop application.`);
    }

    try {
        if (process.platform == "win32") {
            const cmdLine = `tskill ${processName}`;
            await _executeCommandLine(cmdLine);
        } else {
            const pid = await _getProcessId(processName);
            if (pid != undefined) {
                process.kill(pid);
            }
        }
    } catch (err) {
        console.log(`Unable to kill ${application} process. ${err}`);
    }

}

async function _getProcessId(processName: string): Promise<number> {
    return new Promise<number>(async function (resolve, reject) {
        cps.get(function (err: Error, processes: any) {
            try {
                const p = processes.filter(function (p: any) {
                    return (p.name.indexOf(processName) > 0);
                });
                resolve(p.length > 0 ? p[0].pid : undefined);
            }
            catch (err) {
                reject(err);
            }
        });
    });
}
