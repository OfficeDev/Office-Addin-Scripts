import { TSESTree } from "@typescript-eslint/experimental-utils";
import { usageDataObject } from "../defaults";

export = {
  name: "no-context-sync-in-loop",
  meta: {
    type: <"problem" | "suggestion" | "layout">"problem",
    messages: {
      loopedSync:
        "Calling context.sync() inside a loop can lead to poor performance",
    },
    docs: {
      description:
        "Calling context.sync() inside of a loop dramatically increases the time the code runs, proportional to the number of iterations.",
      category: <
        "Best Practices" | "Stylistic Issues" | "Variables" | "Possible Errors"
      >"Best Practices",
      recommended: <false | "error" | "warn">false,
      url: "https://docs.microsoft.com/office/dev/add-ins/concepts/correlated-objects-pattern",
    },
    schema: [],
  },
  create: function (context: any) {
    return {
      ":matches(ForStatement, ForInStatement, WhileStatement, DoWhileStatement, ForOfStatement) CallExpression[callee.object.name='context'][callee.property.name='sync']"(
        node: TSESTree.CallExpression
      ) {
        context.report({
          node: node.callee,
          messageId: "loopedSync",
        });
        usageDataObject.reportSuccess("no-context-sync-in-loop", {
          type: "reported",
        });
      },
    };
  },
};
