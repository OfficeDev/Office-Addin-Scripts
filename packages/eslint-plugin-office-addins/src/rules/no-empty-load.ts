import { TSESTree } from "@typescript-eslint/experimental-utils";

export = {
  name: "no-empty-load",
  meta: {
    type: <"problem" | "suggestion" | "layout">"problem",
    messages: {
      emptyLoad:
        "Calling an empty load can slow down your add-in",
    },
    docs: {
      description:
        "Calling an empty load can cause load unneeded data and slow down your add-in",
      category: <
        "Best Practices" | "Stylistic Issues" | "Variables" | "Possible Errors"
      >"Best Practices",
      recommended: <false | "error" | "warn">false,
      url: "https://docs.microsoft.com/en-us/office/dev/add-ins/develop/application-specific-api-model?view=powerpoint-js-1.1#calling-load-without-parameters-not-recommended",
    },
    schema: [],
  },
  create: function (context: any) {
    return {
      "CallExpression[callee.property.name='run'] :matches(ForStatement, ForInStatement, WhileStatement, DoWhileStatement, ForOfStatement) CallExpression[callee.object.name='context'][callee.property.name='sync']"(
        node: TSESTree.CallExpression
      ) {
        context.report({
          node: node.callee,
          messageId: "emptyLoad",
        });
      },
    };
  },
};
