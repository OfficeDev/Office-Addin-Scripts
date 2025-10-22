// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import inquirer from "inquirer";
import { getOfficeAppName, OfficeApp } from "office-addin-manifest";

export async function chooseOfficeApp(apps: OfficeApp[]): Promise<OfficeApp> {
  const questionName = "app";
  const answer = await inquirer.prompt({
    name: questionName,
    type: "list",
    message: "Which Office app?",
    choices: apps
      .map((app) => ({ name: getOfficeAppName(app), value: app }))
      .sort((a, b) => a.name.localeCompare(b.name)),
  });
  const choice: OfficeApp = answer[questionName];
  return choice;
}
