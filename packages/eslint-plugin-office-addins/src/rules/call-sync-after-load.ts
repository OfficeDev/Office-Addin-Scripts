import { TSESTree } from "@typescript-eslint/experimental-utils";
import { Variable } from "@typescript-eslint/experimental-utils/dist/ts-eslint-scope";
import { findPropertiesRead, findReferences, OfficeApiReference } from "../utils";

export = {
  name: "call-sync-after-load",
  meta: {
    type: <"problem" | "suggestion" | "layout">"suggestion",
    messages: {
      callSyncAfterLoad:
        "Call context.sync() after calling load on '{{name}}' property on the '{{loadValue}}' and before reading properties",
    },
    docs: {
      description:
        "Always call load on an object followed by a context.sync() before reading it or one of its properties.",
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

    function findReadLoadSync(): void {
      const needSync: Set<Variable> = new Set<Variable>();
      const needLoadAndSync: Set<Variable> = new Set<Variable>();

      apiReferences.forEach((apiReference) => {
        const operation = apiReference.operation;
        const reference = apiReference.reference;
        const variable = reference.resolved;

        if (operation === "Write" && variable) {
          needLoadAndSync.add(variable);
        }

        if (operation === "Load" && variable) {
          if (needLoadAndSync.has(variable)) {
            needLoadAndSync.delete(variable);
            needSync.add(variable);
          }
        }

        if (operation === "Sync") {
          needSync.clear();
        }


        const propertyName: string | undefined = findPropertiesRead(
          reference.identifier.parent
        );


        if (
          operation === "Read" &&
          variable &&
          (needSync.has(variable) || needLoadAndSync.has(variable))
        ) {
          const node = reference.identifier;
          context.report({
            node: node,
            messageId: "callSyncAfterLoad",
            data: { name: node.name, loadValue: propertyName },
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
        findReadLoadSync();
      },
    };
  },
};
