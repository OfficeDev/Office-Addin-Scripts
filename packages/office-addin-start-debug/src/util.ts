// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as fetch from 'node-fetch';

export function delay(ms: number) {
    return new Promise(resolve => {
        setTimeout(resolve, ms);
    });
}

//Try to fetch a URL certain amount of times.
export function retryFetch(url: string, fetchOptions = {}, retries: number, retryDelay: number): Promise<fetch.Response> {
    return new Promise<fetch.Response>((resolve, reject) => {
        const wrapper = (n: number) => {
            fetch
                .default(url, fetchOptions)
                .then((res: fetch.Response) => {
                    resolve(res);
                })
                .catch(async err => {
                    if (n > 0) {
                        await delay(retryDelay);
                        wrapper(--n);
                    } else {
                        reject(err);
                    }
                });
        };

        wrapper(retries);
    });
}
