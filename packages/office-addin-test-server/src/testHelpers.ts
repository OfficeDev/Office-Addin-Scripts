const XMLHttpRequest = require("xmlhttprequest").XMLHttpRequest;

const defaultPort: number = 8080;

export async function pingTestServer(port: number = defaultPort): Promise<object> {
    return new Promise<object>(async (resolve, reject) => {
        const serverResponse: object = {};
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

export async function sendTestResults(data: object, port: number = defaultPort): Promise<boolean> {
    return new Promise<boolean>(async (resolve, reject) => {
        const json = JSON.stringify(data);
        const xhr = new XMLHttpRequest();
        const url: string = `https://localhost:${port}/results/`;
        const dataUrl: string = `${url}?data=${encodeURIComponent(json)}`;

        xhr.onreadystatechange = () => {
            if (xhr.readyState === 4 && xhr.status === 200) {
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