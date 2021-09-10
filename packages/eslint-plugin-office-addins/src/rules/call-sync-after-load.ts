import { TSESTree } from "@typescript-eslint/experimental-utils";
import { Reference } from "@typescript-eslint/experimental-utils/dist/ts-eslint-scope";
import { getLoadArgument } from "../utils/load";
import {
  findPropertiesRead,
  findOfficeApiReferences,
  OfficeApiReference,
} from "../utils/utils";

export = {
  name: "call-sync-after-load",
  meta: {
    type: <"problem" | "suggestion" | "layout">"suggestion",
    messages: {
      callSyncAfterLoad:
        "Call context.sync() after calling load on '{{name}}' for property '{{loadValue}}' and before reading the property.",
    },
    docs: {
      description:
        "Always call context.sync() between loading one or more properties on objects and reading any of those properties.",
      category: <
        "Best Practices" | "Stylistic Issues" | "Variables" | "Possible Errors"
      >"Best Practices",
      recommended: <false | "error" | "warn">false,
      url: "https://docs.microsoft.com/office/dev/add-ins/develop/application-specific-api-model#load",
    },
    schema: [],
  },
  create: function (context: any) {
    type VariableProperty = {
      variable: string;
      property: string;
    };

    class VariablePropertySet extends Set {
      add(variableProperty: VariableProperty) {
        return super.add(JSON.stringify(variableProperty));
      }
      has(variableProperty: VariableProperty) {
        return super.has(JSON.stringify(variableProperty));
      }
    }

    let apiReferences: OfficeApiReference[] = [];

    function findLoadBeforeSync(): void {
      const needSync: VariablePropertySet = new VariablePropertySet();

      apiReferences.forEach((apiReference) => {
        const operation = apiReference.operation;
        const reference: Reference = apiReference.reference;
        const identifier: TSESTree.Node = reference.identifier;
        const variable = reference.resolved;

        if (
          operation === "Load" &&
          variable &&
          identifier.parent?.type == TSESTree.AST_NODE_TYPES.MemberExpression
        ) {
          const propertyName: string | undefined = getLoadArgument(
            identifier.parent
          );
          if (propertyName) {
            needSync.add({ variable: variable.name, property: propertyName });
          }
        }

        if (operation === "Sync") {
          needSync.clear();
        }

        if (operation === "Read" && variable) {
          const propertyName: string = findPropertiesRead(
            reference.identifier.parent
          );

          if (
            needSync.has({ variable: variable.name, property: propertyName })
          ) {
            const node = reference.identifier;
            context.report({
              node: node,
              messageId: "callSyncAfterLoad",
              data: { name: node.name, loadValue: propertyName },
            });
          }
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
        findLoadBeforeSync();
      },
    };
  },
};
