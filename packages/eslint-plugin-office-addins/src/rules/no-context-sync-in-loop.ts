import { ESLintUtils, TSESTree } from "@typescript-eslint/utils";

export default ESLintUtils.RuleCreator(
  () =>
    "https://docs.microsoft.com/office/dev/add-ins/concepts/correlated-objects-pattern",
)({
  name: "no-context-sync-in-loop",
  meta: {
    type: "problem",
    messages: {
      loopedSync:
        "Calling context.sync() inside a loop can lead to poor performance",
    },
    docs: {
      description:
        "Calling context.sync() inside of a loop dramatically increases the time the code runs, proportional to the number of iterations.",
    },
    schema: [],
  },
  create: function (context) {
    return {
      ":matches(ForStatement, ForInStatement, WhileStatement, DoWhileStatement, ForOfStatement) CallExpression[callee.object.name='context'][callee.property.name='sync']"(
        node: TSESTree.CallExpression,
      ) {
        context.report({
          node: node.callee,
          messageId: "loopedSync",
        });
      },
    };
  },
  defaultOptions: [],
});
