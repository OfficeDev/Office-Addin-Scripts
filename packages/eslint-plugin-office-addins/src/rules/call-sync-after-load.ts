import { ESLintUtils, TSESTree } from "@typescript-eslint/utils";
import { Reference } from "@typescript-eslint/scope-manager";
import { isLoadCall, parsePropertiesArgument } from "../utils/load";
import {
  findPropertiesRead,
  findOfficeApiReferences,
  OfficeApiReference,
  findCallExpression,
} from "../utils/utils";

export default ESLintUtils.RuleCreator(
  () =>
    "https://docs.microsoft.com/office/dev/add-ins/develop/application-specific-api-model#load",
)({
  name: "call-sync-after-load",
  meta: {
    type: "suggestion",
    messages: {
      callSyncAfterLoad:
        "Call context.sync() after calling load on '{{name}}' for the property '{{loadValue}}' and before reading the property.",
    },
    docs: {
      description:
        "Always call context.sync() between loading one or more properties on objects and reading any of those properties.",
    },
    schema: [],
  },
  create: function (context: any) {
    const sourceCode = context.sourceCode ?? context.getSourceCode();
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
          const propertiesArgument = getPropertiesArgument(identifier);
          const propertyNames: string[] = propertiesArgument
            ? parsePropertiesArgument(propertiesArgument)
            : ["*"];
          propertyNames.forEach((propertyName: string) => {
            needSync.add({
              variable: variable.name,
              property: propertyName,
            });
          });
        } else if (operation === "Sync") {
          needSync.clear();
        } else if (operation === "Read" && variable) {
          const propertyName: string = findPropertiesRead(
            reference.identifier.parent,
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

    function getPropertiesArgument(
      identifier: TSESTree.Identifier | TSESTree.JSXIdentifier,
    ): TSESTree.CallExpressionArgument | undefined {
      let propertiesArgument;
      if (
        identifier.parent?.type === TSESTree.AST_NODE_TYPES.MemberExpression
      ) {
        // Look for <obj>.load(...) call
        const methodCall = findCallExpression(identifier.parent);

        if (methodCall && isLoadCall(methodCall)) {
          propertiesArgument = methodCall.arguments[0];
        }
      } else if (
        identifier.parent?.type === TSESTree.AST_NODE_TYPES.CallExpression
      ) {
        // Look for context.load(<obj>, "...") call
        const args: TSESTree.CallExpressionArgument[] =
          identifier.parent.arguments;
        if (
          isLoadCall(identifier.parent) &&
          args[0] == identifier &&
          args.length < 3
        ) {
          propertiesArgument = args[1];
        }
      }

      return propertiesArgument;
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
        findLoadBeforeSync();
      },
    };
  },
  defaultOptions: [],
});
