import { TSESTree } from "@typescript-eslint/experimental-utils";

export = {
  name: "load-sync-read",
  meta: {
    type: <"problem" | "suggestion" | "layout">"suggestion",
    messages: {
      loadSyncRead:
        "Call load on '{{name}}' for '{{loadValue}}' followed by context.sync() before reading the object or its properties",
    },
    docs: {
      description: "Always call load on an object followed by a sync before reading it or one of its properties.",
      category: <
        "Best Practices" | "Stylistic Issues" | "Variables" | "Possible Errors"
      >"Best Practices",
      recommended: <false | "error" | "warn">false,
      url: "https://docs.microsoft.com/en-us/office/dev/add-ins/develop/application-specific-api-model?view=powerpoint-js-1.1#load",
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
          messageId: "loadSyncRead",
        });
      },
    };
  },
};
