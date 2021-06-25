import {
  AST_NODE_TYPES,
  TSESTree,
} from "@typescript-eslint/experimental-utils";
import {
  Reference,
  Scope,
  Variable,
} from "@typescript-eslint/experimental-utils/dist/ts-eslint-scope";
import { isGetFunction, isLoadFunction } from "../utils";

export = {
  name: "call-sync-before-read",
  meta: {
    type: <"problem" | "suggestion" | "layout">"problem",
    messages: {
      callSync: "Call sync before trying to read '{{name}}'.",
    },
    docs: {
      description: "",
      category: <
        "Best Practices" | "Stylistic Issues" | "Variables" | "Possible Errors"
      >"Possible Errors",
      recommended: <false | "error" | "warn">false,
      url: "https://docs.microsoft.com/en-us/office/dev/add-ins/develop/application-specific-api-model?view=powerpoint-js-1.1#sync",
    },
    schema: [],
  },
  create: function (context: any) {
    const apiReferences: OfficeApiReference[] = [];
    const proxyVariables: Set<Variable> = new Set<Variable>();

    type OfficeApiReference = {
      operation: "Read" | "Load" | "Write" | "Sync";
      reference: Reference;
    };

    function isContextSyncIdentifier(node: TSESTree.Identifier): boolean {
      return (
        node.name === "context" &&
        node.parent?.type === AST_NODE_TYPES.MemberExpression &&
        node.parent?.parent?.type === AST_NODE_TYPES.CallExpression &&
        node.parent?.property.type === AST_NODE_TYPES.Identifier &&
        node.parent?.property.name === "sync"
      );
    }

    function isLoadReference(node: TSESTree.Identifier) {
      return (
        node.parent &&
        node.parent.type === AST_NODE_TYPES.MemberExpression &&
        isLoadFunction(node.parent)
      );
    }

    function findReferences(scope: Scope): void {
      scope.references.forEach((reference) => {
        if (
          reference.isWrite() &&
          reference.writeExpr &&
          isGetFunction(reference.writeExpr) &&
          reference.resolved
        ) {
          proxyVariables.add(reference.resolved);
          apiReferences.push({ operation: "Write", reference: reference });
        } else if (isContextSyncIdentifier(reference.identifier)) {
          apiReferences.push({ operation: "Sync", reference: reference });
        } else if (
          reference.isRead() &&
          reference.resolved &&
          proxyVariables.has(reference.resolved)
        ) {
          if (isLoadReference(reference.identifier)) {
            apiReferences.push({ operation: "Load", reference: reference });
          } else {
            apiReferences.push({ operation: "Read", reference: reference });
          }
        }
      });

      scope.childScopes.forEach(findReferences);
    }

    function findReadBeforeSync(): void {
      const needSync: Set<Variable> = new Set<Variable>();

      apiReferences.forEach((apiReference) => {
        const operation = apiReference.operation;
        const reference = apiReference.reference;

        if (operation === "Write" && reference.resolved) {
          needSync.add(reference.resolved);
        }

        if (operation === "Sync") {
          needSync.clear();
        }

        if (
          operation === "Read" &&
          reference.resolved &&
          needSync.has(reference.resolved)
        ) {
          const node = reference.identifier;
          context.report({
            node: node,
            messageId: "callSync",
            data: { name: node.name },
          });
        }
      });
    }

    return {
      Program(
        programNode: TSESTree.Node /* eslint-disable-line no-unused-vars */
      ) {
        findReferences(context.getScope());
        apiReferences.sort((left, right) => {
          return (
            left.reference.identifier.range[1] - right.reference.identifier.range[1]
          );
        });
        findReadBeforeSync();
      },
    };
  },
};
