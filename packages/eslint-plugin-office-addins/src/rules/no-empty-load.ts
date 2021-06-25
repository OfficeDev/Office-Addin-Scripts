import { TSESTree } from "@typescript-eslint/experimental-utils";
import {
  Reference,
  Scope,
  Variable,
} from "@typescript-eslint/experimental-utils/dist/ts-eslint-scope";
import { isGetFunction, isLoadFunction } from "../utils";

export = {
  name: "no-empty-load",
  meta: {
    type: <"problem" | "suggestion" | "layout">"problem",
    messages: {
      emptyLoad: "Calling load without any argument can slow down your add-in",
    },
    docs: {
      description:
        "Calling load without any argument can cause load unneeded data and slow down your add-in",
      category: <
        "Best Practices" | "Stylistic Issues" | "Variables" | "Possible Errors"
      >"Best Practices",
      recommended: <false | "error" | "warn">false,
      url: "https://docs.microsoft.com/en-us/office/dev/add-ins/develop/application-specific-api-model?view=powerpoint-js-1.1#calling-load-without-parameters-not-recommended",
    },
    schema: [],
  },
  create: function (context: any) {
    function isEmptyLoad(node: TSESTree.MemberExpression): boolean {
      return (
        isLoadFunction(node) &&
        node.parent?.type == TSESTree.AST_NODE_TYPES.CallExpression &&
        node.parent.arguments.length === 0
      );
    }

    function findEmptyLoad(scope: Scope) {
      scope.variables.forEach((variable: Variable) => {
        let getFound: boolean = false;
        variable.references.forEach((reference: Reference) => {
          const node: TSESTree.Node = reference.identifier;

          if (
            reference.isWrite() &&
            reference.writeExpr &&
            isGetFunction(reference.writeExpr) &&
            reference.resolved
          ) {
            getFound = false; // In case of reassignment
            if (
              reference.writeExpr &&
              reference.resolved &&
              isGetFunction(reference.writeExpr)
            ) {
              getFound = true;
              return;
            }
          }

          if (!getFound) {
            // If reference was not related to a previous get
            return;
          }

          if (
            node.parent?.type === TSESTree.AST_NODE_TYPES.MemberExpression &&
            isEmptyLoad(node.parent)
          ) {
            context.report({
              node: node.parent,
              messageId: "emptyLoad",
            });
          }
        });
      });
      scope.childScopes.forEach(findEmptyLoad);
    }

    return {
      Program() {
        findEmptyLoad(context.getScope());
      },
    };
  },
};
