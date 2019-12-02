// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.
/*
    This file provides the provides functionality to get Microsoft Graph data.
*/

import { ODataHelper } from './odata-helper';
import { showMessage } from './message-helper';
import * as $ from 'jquery';

export class MSGraphHelper {

    private static domain: string = 'graph.microsoft.com';
    private static versionURLsegment: string = '/v1.0';

    static getGraphData(accessToken: string, apiURLsegment: string, queryParamsSegment?: string) {
        return new Promise<any>(async (resolve, reject) => {
            try {
                const oData = await ODataHelper.getData(accessToken, this.domain, apiURLsegment, this.versionURLsegment, queryParamsSegment);
                resolve(oData);
            } catch (err) {
                reject(`Error get Graph data. \n${err}`);
            }
        });
    }

    static async getGraphToken(bootstrapToken): Promise<any> {
        try {
            let response = await $.ajax({
                type: "GET",
                url: "/auth",
                headers: { Authorization: "Bearer " + bootstrapToken },
                cache: false
            });
            return response;
        } catch (err) {
            throw new Error(`Error getting Graph token. \n${err}`);
        }

    }

    static async makeGraphApiCall(accessToken: string): Promise<any> {
        try {
            const response = await $.ajax({
                type: "GET",
                url: "/getuserdata",
                headers: { access_token: accessToken },
                cache: false
            })
            return response;
        } catch (err) {
            showMessage(`Error from Microsoft Graph. \n${err}`);
        }
    }
}
