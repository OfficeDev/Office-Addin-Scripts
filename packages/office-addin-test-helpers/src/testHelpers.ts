// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.
import * as defaults from "./defaults";
import * as fetch from "isomorphic-fetch";

/**
 * Ping the test server
 * @param port Specifies port to ping
 * @param useHttps Specifies protocol to use when pinging port (https or http)
 */
export async function pingTestServer(port: number = defaults.httpsPort, useHttps: boolean = true): Promise<object> {
    return new Promise<object>(async (resolve, reject) => {
        const serverResponse: any = {};
        try {
            const protocol = useHttps ? "https" : "http";
            const pingUrl: string = `${protocol}://localhost:${port}/ping`
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

/**
 * Ping the test server
 * @param data Specifies data object containing test results
 * @param port Specifies port to send data to
 */
export async function sendTestResults(data: object, port: number = defaults.httpsPort): Promise<boolean> {
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


