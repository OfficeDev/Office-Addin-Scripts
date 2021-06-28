import { TSESTree } from "@typescript-eslint/experimental-utils";
import { Variable } from "@typescript-eslint/experimental-utils/dist/ts-eslint-scope";
import { findReferences, OfficeApiReference } from "../utils";

export = {
  name: "load-sync-read",
  meta: {
    type: <"problem" | "suggestion" | "layout">"suggestion",
    messages: {
      loadSyncRead:
        "Call load on '{{name}}' for '{{loadValue}}' followed by context.sync() before reading the object or its properties",
    },
    docs: {
      description:
        "Always call load on an object followed by a sync before reading it or one of its properties.",
      category: <
        "Best Practices" | "Stylistic Issues" | "Variables" | "Possible Errors"
      >"Best Practices",
      recommended: <false | "error" | "warn">false,
      url: "https://docs.microsoft.com/en-us/office/dev/add-ins/develop/application-specific-api-model?view=powerpoint-js-1.1#load",
    },
    schema: [],
  },
  create: function (context: any) {
    let apiReferences: OfficeApiReference[] = [];

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
            messageId: "loadSyncRead",
            data: { name: node.name },
          });
        }
      });
    }

    return {
      Program(
        programNode: TSESTree.Node /* eslint-disable-line no-unused-vars */
      ) {
        apiReferences = findReferences(context.getScope());
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
};
