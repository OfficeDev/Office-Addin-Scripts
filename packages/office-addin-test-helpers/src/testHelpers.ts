// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.
import fetch from "isomorphic-fetch";
export const defaultPort: number = 4201;

export interface TestServerResponse {
  status: number;
  platform: string;
  error: any;
}

export async function pingTestServer(port: number = defaultPort): Promise<TestServerResponse> {
  const serverResponse: TestServerResponse = { status: 0, platform: "", error: null };
  try {
    const pingUrl: string = `https://localhost:${port}/ping`;
    const response = await fetch(pingUrl);
    serverResponse.status = response.status;
    const text = await response.text();
    serverResponse.platform = text;
    return Promise.resolve(serverResponse);
  } catch (err) {
    serverResponse.error = err;
    return Promise.reject(serverResponse);
  }
}

export async function sendTestResults(data: object, port: number = defaultPort): Promise<boolean> {
  const json = JSON.stringify(data);
  const url: string = `https://localhost:${port}/results/`;
  const dataUrl: string = url + "?data=" + encodeURIComponent(json);

  try {
    await fetch(dataUrl, {
      method: "post",
      body: JSON.stringify(data),
      headers: { "Content-Type": "application/json" },
    });
    return true;
  } catch {
    return false;
  }
}
