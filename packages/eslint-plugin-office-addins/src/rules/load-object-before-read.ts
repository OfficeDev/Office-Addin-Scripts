import { TSESTree } from "@typescript-eslint/typescript-estree";

export = {
  name: "load-object-before-read",
  meta: {
    type: <"problem" | "suggestion" | "layout">"problem",
    messages: {
      loadBeforeRead: "An explicit load call needs to be made before reading a proxu object",
    },
    docs: {
      description: "Before you can read the properties of a proxy object, you must explicitly load the properties",
      category: <"Best Practices" | "Stylistic Issues" | "Variables" | "Possible Errors">"Possible Errors",
      recommended: <false | "error" | "warn">false,
      url: "https://docs.microsoft.com/en-us/office/dev/add-ins/develop/application-specific-api-model?view=powerpoint-js-1.1#load",
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
          messageId: "loadBeforeRead",
        });
      },
    };
  },
};
