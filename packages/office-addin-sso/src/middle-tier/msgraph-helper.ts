// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.
/*
    This file provides the provides functionality to get Microsoft Graph data.
*/

import { ODataHelper } from "./odata-helper";

const domain: string = "graph.microsoft.com";
const versionURLsegment: string = "/v1.0";

export async function getGraphData(
  accessToken: string,
  apiURLsegment: string,
  queryParamsSegment?: string
): Promise<any> {
  try {
    const oData = await ODataHelper.getData(
      accessToken,
      domain,
      apiURLsegment,
      versionURLsegment,
      queryParamsSegment
    );
    return Promise.resolve(oData);
  } catch (err) {
    return Promise.reject(`Error get Graph data. \n${err}`);
  }
}
