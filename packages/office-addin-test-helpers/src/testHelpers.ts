import * as fetch from "isomorphic-fetch";
export const defaultPort: number = 4201;

export async function pingTestServer(port: number = defaultPort): Promise<object> {
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

export async function sendTestResults(data: object, port: number = defaultPort): Promise<boolean> {
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

