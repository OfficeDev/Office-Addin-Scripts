import { Variable } from "@typescript-eslint/experimental-utils/dist/ts-eslint-scope";
import { findOfficeApiReferences, OfficeApiReference } from "../utils";

export = {
  name: "call-sync-before-read",
  meta: {
    type: <"problem" | "suggestion" | "layout">"problem",
    messages: {
      callSync: "Call sync before trying to read '{{name}}'.",
    },
    docs: {
      description: "Always call load on the object's properties followed by a context.sync() before reading them.",
      category: <
        "Best Practices" | "Stylistic Issues" | "Variables" | "Possible Errors"
      >"Possible Errors",
      recommended: <false | "error" | "warn">false,
      url: "https://docs.microsoft.com/en-us/office/dev/add-ins/develop/application-specific-api-model?view=powerpoint-js-1.1#sync",
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
        const variable = reference.resolved;

        if (operation === "Write" && variable) {
          needSync.add(variable);
        }

        if (operation === "Sync") {
          needSync.clear();
        }

        if (operation === "Read" && variable && needSync.has(variable)) {
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
      Program() {
        apiReferences = findOfficeApiReferences(context.getScope());
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
