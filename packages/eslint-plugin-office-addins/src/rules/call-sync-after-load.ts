import { TSESTree } from "@typescript-eslint/utils";
import { Reference } from "@typescript-eslint/utils/dist/ts-eslint-scope";
import {
  isLoadFunction,
  parseLoadArguments,
  parsePropertiesArgument,
} from "../utils/load";
import {
  findPropertiesRead,
  findOfficeApiReferences,
  OfficeApiReference,
  findTopLevelExpression,
} from "../utils/utils";

export = {
  name: "call-sync-after-load",
  meta: {
    type: <"problem" | "suggestion" | "layout">"suggestion",
    messages: {
      callSyncAfterLoad:
        "Call context.sync() after calling load on '{{name}}' for the property '{{loadValue}}' and before reading the property.",
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

        if (operation === "Load" && variable) {
          if (
            identifier.parent?.type == TSESTree.AST_NODE_TYPES.MemberExpression
          ) {
            // Look for <obj>.load(...) call
            const topParent = findTopLevelExpression(identifier.parent);

            if (
              isLoadFunction(topParent) &&
              topParent.parent?.type === TSESTree.AST_NODE_TYPES.CallExpression
            ) {
              const argument = topParent.parent.arguments[0];
              let propertyNames: string[] = argument
                ? parsePropertiesArgument(argument)
                : ["*"];
              propertyNames.forEach((propertyName: string) => {
                needSync.add({
                  variable: variable.name,
                  property: propertyName,
                });
              });
              return;
            }
          } else if (
            // Look for context.load(<obj>, "...") call
            identifier.parent?.type == TSESTree.AST_NODE_TYPES.CallExpression
          ) {
            const callee: TSESTree.MemberExpression = identifier.parent
              .callee as TSESTree.MemberExpression;
            const args: TSESTree.CallExpressionArgument[] =
              identifier.parent.arguments;
            if (
              isLoadFunction(callee) &&
              args[0] == identifier &&
              args.length < 3
            ) {
              const propertyNames: string[] = args[1]
                ? parsePropertiesArgument(args[1])
                : ["*"];
              propertyNames.forEach((propertyName: string) => {
                needSync.add({
                  variable: variable.name,
                  property: propertyName,
                });
              });
            }
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
            needSync.has({ variable: variable.name, property: propertyName }) ||
            needSync.has({ variable: variable.name, property: "*" })
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
