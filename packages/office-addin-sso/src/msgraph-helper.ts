// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.
/*
    This file provides the provides functionality to get Microsoft Graph data.
*/

import { showMessage } from "./message-helper";
import $ from "jquery";

export async function makeGraphApiCall(bootstrapToken: string): Promise<any> {
  try {
    const response = await $.ajax({
      type: "GET",
      url: "/getuserdata",
      headers: { Authorization: "Bearer " + bootstrapToken },
      cache: false,
    });
    return response;
  } catch (err) {
    showMessage(`Error from Microsoft Graph. \n${err}`);
  }
}
