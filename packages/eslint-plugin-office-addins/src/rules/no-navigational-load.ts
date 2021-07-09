import { TSESTree } from "@typescript-eslint/experimental-utils";
import {
  Reference,
  Scope,
  Variable,
} from "@typescript-eslint/experimental-utils/dist/ts-eslint-scope";
import { isLoadFunction, getLiteralArgumentName } from "../utils/utils";
import { isGetFunction } from "../utils/getFunction";
import { getPropertyType, PropertyType } from "../utils/propertiesType";

export = {
  name: "no-navigational-load",
  meta: {
    type: <"problem" | "suggestion" | "layout">"problem",
    messages: {
      navigationalLoad:
        "Calling load on the navigation property '{{loadValue}}' slows down your add-in.",
    },
    docs: {
      description:
        "Calling load on a navigation property causes unneeded data to load and slows down your add-in.",
      category: <
        "Best Practices" | "Stylistic Issues" | "Variables" | "Possible Errors"
      >"Best Practices",
      recommended: <false | "error" | "warn">false,
      url: "https://docs.microsoft.com/office/dev/add-ins/develop/application-specific-api-model#scalar-and-navigation-properties",
    },
    schema: [],
  },
  create: function (context: any) {
    function trimPropertyName(propertyName: string): string {
      return propertyName.split("/").pop() ?? "";
    }

    function isNotScalarLoad(
      node: TSESTree.MemberExpression,
      propertyName: string
    ): boolean {
      const propertyType: PropertyType = getPropertyType(
        trimPropertyName(propertyName)
      );

      return isLoadFunction(node) && propertyType !== PropertyType.scalar;
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
            const propertyName: string = getLiteralArgumentName(
              reference.identifier.parent
            );
            if (isNotScalarLoad(node.parent, propertyName)) {
              context.report({
                node: node.parent,
                messageId: "navigationalLoad",
                data: { name: node.name, loadValue: propertyName },
              });
            }
          }
        });
      });
      scope.childScopes.forEach(findNavigationalLoad);
    }

    return {
      Program() {
        findNavigationalLoad(context.getScope());
      },
    };
  },
};
