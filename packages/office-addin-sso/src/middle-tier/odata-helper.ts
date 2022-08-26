// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.
/*
    This file provides the provides functionality to get data from OData-compliant endpoints.
*/

import * as https from "https";

export class ODataHelper {
  static getData(
    accessToken: string,
    domain: string,
    apiURLsegment: string,
    apiVersion?: string,
    // If any part of queryParamsSegment comes from user input,
    // be sure that it is sanitized so that it cannot be used in
    // a Response header injection attack.
    queryParamsSegment?: string
  ) {
    return new Promise<any>((resolve, reject) => {
      const options = {
        host: domain,
        path: apiVersion + apiURLsegment + queryParamsSegment,
        method: "GET",
        headers: {
          "Content-Type": "application/json",
          Accept: "application/json",
          Authorization: "Bearer " + accessToken,
          "Cache-Control": "private, no-cache, no-store, must-revalidate",
          Expires: "-1",
          Pragma: "no-cache",
        },
      };

      https
        .get(options, (response) => {
          let body = "";
          response.on("data", (d) => {
            body += d;
          });
          response.on("end", () => {
            // The response from the OData endpoint might be an error, say a
            // 401 if the endpoint requires an access token and it was invalid
            // or expired. But a message is not an error in the call of https.get,
            // so the "on('error', reject)" line below isn't triggered.
            // So, the code distinguishes success (200) messages from error
            // messages and sends a JSON object to the caller with either the
            // requested OData or error information.

            let error;
            if (response.statusCode === 200) {
              let parsedBody = JSON.parse(body);
              resolve(parsedBody);
            } else {
              error = new Error();
              error.code = response.statusCode;
              error.message = response.statusMessage;

              // The error body sometimes includes an empty space
              // before the first character, remove it or it causes an error.
              body = body.trim();
              error.bodyCode = JSON.parse(body).error.code;
              error.bodyMessage = JSON.parse(body).error.message;
              resolve(error);
            }
          });
        })
        .on("error", reject);
    });
  }
}
