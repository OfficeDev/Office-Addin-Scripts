/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global $ */

export function showMessage(text: string): void {
  $(".welcome-body").hide();
  $("#message-area").show();
  $("#message-area").text(text);
}
