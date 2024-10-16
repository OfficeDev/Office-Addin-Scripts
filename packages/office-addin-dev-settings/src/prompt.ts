// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import inquirer from "inquirer";
import { getOfficeAppName, OfficeApp } from "office-addin-manifest";

export async function chooseOfficeApp(apps: OfficeApp[]): Promise<OfficeApp> {
  const questionName = "app";
  const question: inquirer.ListQuestionOptions<inquirer.Answers> = {
    choices: apps
      .map((app) => {
        return { name: getOfficeAppName(app), value: app };
      })
      .sort((first, second) => first.name.localeCompare(second.name)),
    message: "Which Office app?",
    name: questionName,
    type: "list",
  };

  const answer = await inquirer.prompt([question]);
  const choice: OfficeApp = answer[questionName];
  return choice;
}
