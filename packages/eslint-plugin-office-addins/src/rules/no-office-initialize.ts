import { TSESTree } from "@typescript-eslint/experimental-utils";
import { usageDataObject } from "../defaults";

export = {
  name: "no-office-initialize",
  meta: {
    type: <"problem" | "suggestion" | "layout">"suggestion",
    messages: {
      noOfficeInitialize:
        "Office.onReady() is preferred over Office.initialize.",
    },
    docs: {
      description: "Office.onReady() is more flexible than Office.initialize.",
      category: <
        "Best Practices" | "Stylistic Issues" | "Variables" | "Possible Errors"
      >"Best Practices",
      recommended: <false | "error" | "warn">false,
      url: "https://docs.microsoft.com/office/dev/add-ins/develop/initialize-add-in#initialize-with-officeonready",
    },
    schema: [],
  },
  create: function (context: any) {
    return {
      "AssignmentExpression[left.object.name='Office'][left.property.name='initialize']"(
        node: TSESTree.AssignmentExpression
      ) {
        context.report({
          node: node,
          messageId: "noOfficeInitialize",
        });
        usageDataObject.reportSuccess("no-office-initialize");
      },
    };
  },
};
