/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in root of repo. -->
 *
 * This file defines the routes within the authRoute router.
 */

import * as express from 'express';
import * as fetch from 'node-fetch';
import * as form from 'form-urlencoded';
import { sendUsageDataException, sendUsageDataSuccessEvent } from './defaults';

export class AuthRouter {
    public router: express.Router;
    constructor() {
        this.router = express.Router();

        this.router.get('/', async function (req: any, res: any, next: any) {
            const authorization = req.get('Authorization');
            const scopeName: string = process.env.SCOPE || 'User.Read'
            if (authorization == null) {
                let error = new Error('No Authorization header was found.');
                next(error);
            }
            else {
                const [/* schema */, jwt] = authorization.split(' ');
                const formParams = {
                    client_id: process.env.CLIENT_ID,
                    client_secret: process.env.secret,
                    grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
                    assertion: jwt,
                    requested_token_use: 'on_behalf_of',
                    scope: [scopeName].join(' ')
                };

                const stsDomain: string = 'https://login.microsoftonline.com';
                const tenant: string = 'common';
                const tokenURLSegment: string = 'oauth2/v2.0/token';

                try {
                    const tokenResponse = await fetch(`${stsDomain}/${tenant}/${tokenURLSegment}`, {
                        method: 'POST',
                        body: form.default(formParams),
                        headers: {
                            'Accept': 'application/json',
                            'Content-Type': 'application/x-www-form-urlencoded'
                        }
                    });
                    const json = await tokenResponse.json();
                    res.send(json);

                    // Send usage data
                    sendUsageDataSuccessEvent('authRouter', {scope: scopeName});
                }
                catch (error) {
                    res.status(500).send(error);
                    sendUsageDataException('authRouter', error);
                }
            }
        });
    }
}

