// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.
/*
    This file provides the provides functionality to get Microsoft Graph data.
*/

import { ODataHelper } from "./odata-helper";
import { showMessage } from "./message-helper";
import * as $ from "jquery";

const domain: string = "graph.microsoft.com";
const versionURLsegment: string = "/v1.0";

export async function getGraphData(
  accessToken: string,
  apiURLsegment: string,
  queryParamsSegment?: string
): Promise<any> {
  return new Promise<any>(async (resolve, reject) => {
    try {
      const oData = await ODataHelper.getData(
        accessToken,
        domain,
        apiURLsegment,
        versionURLsegment,
        queryParamsSegment
      );
      resolve(oData);
    } catch (err) {
      reject(`Error get Graph data. \n${err}`);
    }
  });
}

export async function getGraphToken(bootstrapToken): Promise<any> {
  try {
    let response = await $.ajax({
      type: "GET",
      url: "/auth",
      headers: { Authorization: "Bearer " + bootstrapToken },
      cache: false,
    });
    return response;
  } catch (err) {
    throw new Error(`Error getting Graph token. \n${err}`);
  }
}

export async function makeGraphApiCall(accessToken: string): Promise<any> {
  try {
    const response = await $.ajax({
      type: "GET",
      url: "/getuserdata",
      headers: { access_token: accessToken },
      cache: false,
    });
    return response;
  } catch (err) {
    showMessage(`Error from Microsoft Graph. \n${err}`);
  }
}
