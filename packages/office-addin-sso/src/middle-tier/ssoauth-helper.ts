/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in root of repo. -->
 *
 * This file defines the routes within the authRoute router.
 */

import fetch from "node-fetch";
import * as form from "form-urlencoded";
import * as jwt from "jsonwebtoken";
import { JwksClient } from "jwks-rsa";

/* global process, console */

const DISCOVERY_KEYS_ENDPOINT = "https://login.microsoftonline.com/common/discovery/v2.0/keys";

export async function getAccessToken(authorization: string): Promise<any> {
  const scopeName: string = process.env.SCOPE || "User.Read";
  if (authorization === null) {
    let error = new Error("No Authorization header was found.");
    Promise.reject(error);
  } else {
    const [, /* schema */ jwt] = authorization.split(" ");
    const formParams = {
      client_id: process.env.CLIENT_ID,
      client_secret: process.env.secret,
      grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
      assertion: jwt,
      requested_token_use: "on_behalf_of",
      scope: [scopeName].join(" "),
    };

    const stsDomain: string = "https://login.microsoftonline.com";
    const tenant: string = "common";
    const tokenURLSegment: string = "oauth2/v2.0/token";

    try {
      const tokenResponse = await fetch(`${stsDomain}/${tenant}/${tokenURLSegment}`, {
        method: "POST",
        body: form.default(formParams),
        headers: {
          Accept: "application/json",
          "Content-Type": "application/x-www-form-urlencoded",
        },
      });
      const json = await tokenResponse.json();
      return json;
    } catch (error) {
      Promise.reject(error);
    }
  }
}

export function validateJwt(req, res, next): void {
  const authHeader = req.headers.authorization;
  if (authHeader) {
    const token = authHeader.split(" ")[1];

    const validationOptions = {
      audience: process.env.CLIENT_ID,
    };

    jwt.verify(token, getSigningKeys, validationOptions, (err) => {
      //custom logic to regex search for tenant id in the issuer.
      //test multi tenant setup.
      //test msa

      if (err) {
        console.log(err);
        return res.sendStatus(403);
      }

      next();
    });
  }
}

function getSigningKeys(header: any, callback: any) {
  var client: JwksClient = new JwksClient({
    jwksUri: DISCOVERY_KEYS_ENDPOINT,
  });

  client.getSigningKey(header.kid, function (err, key) {
    callback(null, key.getPublicKey());
  });
}
