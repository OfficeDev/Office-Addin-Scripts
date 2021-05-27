import { TSESTree } from "@typescript-eslint/typescript-estree";

export = {
  name: "no-context-sync-in-loop",
  meta: {
    type: <"problem" | "suggestion" | "layout">"problem",
    messages: {
      loopedSync:
        "Calling context.sync() inside a loop can lead to bad performance",
    },
    docs: {
      description:
        "Calling context.sync() inside of a loop dramatically increates the time the code runs the more iterations that are run",
      category: <
        "Best Practices" | "Stylistic Issues" | "Variables" | "Possible Errors"
      >"Best Practices",
      recommended: <false | "error" | "warn">false,
      url:
        "https://docs.microsoft.com/en-us/office/dev/add-ins/concepts/correlated-objects-pattern?view=powerpoint-js-1.1",
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
          messageId: "loopedSync",
        });
      },
    };
  },
};
