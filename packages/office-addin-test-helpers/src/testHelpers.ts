
import * as childProcess from "child_process";
import * as cps from "current-processes";
import * as fetch from "isomorphic-fetch";
let applicationSideloaded: boolean = false;
let devServerProcess: any;
let devServerStarted: boolean = false;

export async function pingTestServer(port: number = 4201): Promise<object> {
    return new Promise<object>(async (resolve, reject) => {
        const serverResponse: any = {};
        try {
            const pingUrl: string = `https://localhost:${port}/ping`;
            const response = await fetch(pingUrl);
            serverResponse["status"] = response.status;
            const text = await response.text();
            serverResponse["platform"] = text;
            resolve(serverResponse);
        } catch (err) {
            serverResponse["status"] = err;
            reject(serverResponse);
        }
    });
}

export async function sendTestResults(data: object, port: number = 4201): Promise<boolean> {
    return new Promise<boolean>(async (resolve, reject) => {
        const json = JSON.stringify(data);
        const url: string = `https://localhost:${port}/results/`;
        const dataUrl: string = url + "?data=" + encodeURIComponent(json);

        try {
            fetch(dataUrl, {
                method: 'post',
                body: JSON.stringify(data),
                headers: { 'Content-Type': 'application/json' },
            });
            resolve(true);
        } catch (err) {
            reject(false);
        }
    });
}


export async function sideloadDesktopApp(application: string, manifestPath: string): Promise<boolean> {
    return new Promise<boolean>(async function (resolve, reject) {
        if (process.platform !== 'win32' && process.platform !== 'darwin') {
            resolve(false);
        }

        try {
            const cmdLine = `node ${process.cwd()}/node_modules/office-toolbox/app/office-toolbox.js sideload -m ${manifestPath} -a ${application}`;
            applicationSideloaded = await executeCommandLine(cmdLine);
            resolve(applicationSideloaded);
        } catch (err) {
            reject(err);
        }
    });
}

export async function startDevServer(): Promise<boolean> {
    return new Promise<boolean>((resolve, reject) => {
        devServerStarted = false;
        const cmdLine = "npm run dev-server";
        devServerProcess = childProcess.spawn(cmdLine, [], {
            detached: true,
            shell: true,
            stdio: "ignore"
        });
        devServerProcess.on("error", (err) => {
            reject(err);
        });

        devServerStarted = devServerProcess.pid != undefined;
        resolve(devServerStarted);
    });
}

async function executeCommandLine(cmdLine: string): Promise<boolean> {
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
export async function teardownTestEnvironment(application: string): Promise<boolean> {
    return new Promise<boolean>(async function (resolve, reject) {
        let applicationKilled: boolean = false;
        let devServerKilled: boolean = false;

        if (applicationSideloaded) {
            try {
                applicationKilled = await closeDesktopApplication(application);
            } catch (err) {
                reject(`Test teardown failed: ${err}`);
            }            
        }

        // if the dev-server was started, kill the spawned process
        if (devServerStarted) {
            try {
                if (process.platform == "win32") {
                    childProcess.spawn("taskkill", ["/pid", devServerProcess.pid, '/f', '/t']);
                } else {
                    devServerProcess.kill();
                }                
                devServerKilled = true;
            } catch (err) {
                reject(`Test teardown failed: ${err}`);
            }
        }
        resolve(applicationKilled && devServerKilled);
    });
}

// Office-JS close workbook API is coming soon.  Once it's available we can stript out this code to kill the Excel process
async function closeDesktopApplication(application: string): Promise<boolean> {
    return new Promise<boolean>(async function (resolve, reject) {
        let processName: string = "";
        switch (application.toLowerCase()) {
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
                reject(`${application} is not a valid Office desktop application.`);
            }

        try {
            let appClosed: boolean = false;
            if (process.platform == "win32") {
                const cmdLine = `tskill ${processName}`;
                appClosed = await executeCommandLine(cmdLine);
            } else {
                const pid = await getProcessId(processName);
                if (pid != undefined) {
                    process.kill(pid);
                    appClosed = true;
                } else {
                    resolve(false);
                }               
            }
            resolve(appClosed);
        } catch (err) {
            reject(`Unable to kill ${application} process. ${err}`);
        }
    });
}

async function getProcessId(processName: string): Promise<number> {
    return new Promise<number>(async function (resolve, reject) {
        cps.get(function (err: Error, processes: any) {
            try {
                const processArray = processes.filter(function (p: any) {
                    return (p.name.indexOf(processName) > 0);
                });
                resolve(processArray.length > 0 ? processArray[0].pid : undefined);
            }
            catch (err) {
                reject(err);
            }
        });
    });
}
