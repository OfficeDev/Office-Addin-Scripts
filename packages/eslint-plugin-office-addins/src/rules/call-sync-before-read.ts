import { ESLintUtils, TSESTree } from "@typescript-eslint/utils";
import { Variable } from "@typescript-eslint/scope-manager";
import { findTopMemberExpression } from "../utils/utils";
import { findOfficeApiReferences, OfficeApiReference } from "../utils/utils";

export default ESLintUtils.RuleCreator(
  () =>
    "https://docs.microsoft.com/office/dev/add-ins/develop/application-specific-api-model#sync",
)({
  name: "call-sync-before-read",
  meta: {
    type: "problem",
    messages: {
      callSync: "Call context.sync() before trying to read '{{name}}'.",
    },
    docs: {
      description:
        "Always call load on the object's properties followed by a context.sync() before reading them.",
    },
    schema: [],
  },
  create: function (context: any) {
    const sourceCode = context.sourceCode ?? context.getSourceCode();
    let apiReferences: OfficeApiReference[] = [];

    function checkPropertyIsRead(node: TSESTree.MemberExpression): boolean {
      const topExpression: TSESTree.MemberExpression =
        findTopMemberExpression(node);
      switch (topExpression.parent?.type) {
        case TSESTree.AST_NODE_TYPES.AssignmentExpression:
          return topExpression.parent.right === topExpression;
        default:
          return true;
      }
    }

    function findReadBeforeSync(): void {
      const needSync: Set<Variable> = new Set<Variable>();

      apiReferences.forEach((apiReference) => {
        const operation = apiReference.operation;
        const reference = apiReference.reference;
        const variable = reference.resolved;

        if (operation === "Get" && variable) {
          needSync.add(variable);
        }

        if (operation === "Sync") {
          needSync.clear();
        }

        if (operation === "Read" && variable && needSync.has(variable)) {
          const node: TSESTree.Node = reference.identifier;
          if (
            node.parent?.type === TSESTree.AST_NODE_TYPES.MemberExpression &&
            checkPropertyIsRead(node.parent)
          ) {
            context.report({
              node: node,
              messageId: "callSync",
              data: { name: node.name },
            });
          }
        }
      });
    }

    return {
      Program(node) {
        const scope = sourceCode.getScope
                    ? sourceCode.getScope(node)
                    : context.getScope();
        apiReferences = findOfficeApiReferences(scope);
        apiReferences.sort((left, right) => {
          return (
            left.reference.identifier.range[1] -
            right.reference.identifier.range[1]
          );
        });
        findReadBeforeSync();
      },
    };
  },
  defaultOptions: [],
});
