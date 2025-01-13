import { ESLintUtils, TSESTree } from "@typescript-eslint/utils";
import { Reference, Scope, Variable } from "@typescript-eslint/scope-manager";
import { isGetFunction } from "../utils/getFunction";
import { parseLoadArguments, isLoadFunction } from "../utils/load";

export default ESLintUtils.RuleCreator(
  () =>
    "https://docs.microsoft.com/office/dev/add-ins/develop/application-specific-api-model#calling-load-without-parameters-not-recommended",
)({
  name: "no-empty-load",
  meta: {
    type: "problem",
    messages: {
      emptyLoad: "Calling load without any argument slows down your add-in.",
    },
    docs: {
      description:
        "Calling load without any argument causes unneeded data to load and slows down your add-in.",
    },
    schema: [],
  },
  create: function (context: any) {
    const sourceCode = context.sourceCode ?? context.getSourceCode();
    function isEmptyLoad(node: TSESTree.MemberExpression): boolean {
      if (isLoadFunction(node)) {
        const propertyNames: string[] = parseLoadArguments(node);
        if (propertyNames.length === 0) {
          return true;
        }

        let foundEmptyProperty = false;
        propertyNames.forEach((property: string) => {
          if (!property) {
            foundEmptyProperty = true;
          }
        });
        return foundEmptyProperty;
      }
      return false;
    }

    function findEmptyLoad(scope: Scope) {
      scope.variables.forEach((variable: Variable) => {
        let getFound: boolean = false;
        variable.references.forEach((reference: Reference) => {
          const node: TSESTree.Node = reference.identifier;

          if (reference.isWrite()) {
            getFound = false; // In case of reassignment
            if (reference.writeExpr && isGetFunction(reference.writeExpr)) {
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
      Program(node) {
        const scope = sourceCode.getScope
                    ? sourceCode.getScope(node)
                    : context.getScope();
        findEmptyLoad(scope);
      },
    };
  },
  defaultOptions: [],
});
