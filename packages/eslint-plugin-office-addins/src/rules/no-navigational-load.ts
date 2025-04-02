import { ESLintUtils, TSESTree } from "@typescript-eslint/utils";
import { Reference, Scope, Variable } from "@typescript-eslint/scope-manager";
import { isGetFunction } from "../utils/getFunction";
import {
  parseLoadArguments,
  isLoadFunction,
  parsePropertiesArgument,
} from "../utils/load";
import { getPropertyType, PropertyType } from "../utils/propertiesType";

export default ESLintUtils.RuleCreator(
  () =>
    "https://docs.microsoft.com/office/dev/add-ins/develop/application-specific-api-model#scalar-and-navigation-properties",
)({
  name: "no-navigational-load",
  meta: {
    type: "problem",
    messages: {
      navigationalLoad:
        "Calling load on the navigation property '{{loadValue}}' slows down your add-in.",
    },
    docs: {
      description:
        "Calling load on a navigation property causes unneeded data to load and slows down your add-in.",
    },
    schema: [],
  },
  create: function (context) {
    const sourceCode = context.sourceCode ?? context.getSourceCode();
    function isLoadingValidPropeties(propertyName: string): boolean {
      const properties = propertyName.split("/");
      const lastProperty = properties.pop();
      if (!lastProperty) return false;

      for (const property of properties) {
        const propertyType = getPropertyType(property);
        if (
          propertyType !== PropertyType.navigational &&
          propertyType !== PropertyType.ambiguous
        ) {
          return false;
        }
      }

      if (lastProperty === "*") {
        return true;
      }
      const propertyType = getPropertyType(lastProperty);
      return (
        propertyType === PropertyType.scalar ||
        propertyType === PropertyType.ambiguous
      );
    }

    function findNavigationalLoad(scope: Scope) {
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
            isLoadFunction(node.parent)
          ) {
            // <obj>.load(...) call
            const propertyNames: string[] = parseLoadArguments(node.parent);
            propertyNames.forEach((propertyName: string) => {
              if (propertyName && !isLoadingValidPropeties(propertyName)) {
                context.report({
                  node: node.parent,
                  messageId: "navigationalLoad",
                  data: { name: node.name, loadValue: propertyName },
                });
              }
            });
          } else if (
            node.parent?.type === TSESTree.AST_NODE_TYPES.CallExpression
          ) {
            //context.load(<obj>, "...") call
            const callee: TSESTree.MemberExpression = node.parent
              .callee as TSESTree.MemberExpression;
            const args: TSESTree.CallExpressionArgument[] =
              node.parent.arguments;
            if (isLoadFunction(callee) && args[0] == node && args.length < 3) {
              const propertyNames: string[] = parsePropertiesArgument(args[1]);
              propertyNames.forEach((propertyName: string) => {
                if (propertyName && !isLoadingValidPropeties(propertyName)) {
                  context.report({
                    node: node.parent,
                    messageId: "navigationalLoad",
                    data: { name: node.name, loadValue: propertyName },
                  });
                }
              });
            }
          }
        });
      });
      scope.childScopes.forEach(findNavigationalLoad);
    }

    return {
      Program(node) {
        const scope = sourceCode.getScope
          ? sourceCode.getScope(node)
          : context.getScope();
        findNavigationalLoad(scope);
      },
    };
  },
  defaultOptions: [],
});
