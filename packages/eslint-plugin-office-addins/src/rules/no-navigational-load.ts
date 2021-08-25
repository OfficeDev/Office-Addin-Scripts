import { TSESTree } from "@typescript-eslint/experimental-utils";
import {
  Reference,
  Scope,
  Variable,
} from "@typescript-eslint/experimental-utils/dist/ts-eslint-scope";
import { isGetFunction } from "../utils/getFunction";
import { getLoadArgument, isLoadFunction } from "../utils/load";
import { getPropertyType, PropertyType } from "../utils/propertiesType";
import { usageDataObject } from "../defaults";

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
    function isLoadingValidPropeties(propertyName: string): boolean {
      const properties = propertyName.split("/");
      const lastProperty = properties.pop();
      if (!lastProperty) return false;

      for (const property of properties) {
        if (getPropertyType(property) !== PropertyType.navigational) {
          return false;
        }
      }

      return getPropertyType(lastProperty) === PropertyType.scalar;
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
            const propertyName: string = getLoadArgument(node.parent);
            if (!isLoadingValidPropeties(propertyName)) {
              context.report({
                node: node.parent,
                messageId: "navigationalLoad",
                data: { name: node.name, loadValue: propertyName },
              });
              usageDataObject.reportSuccess("no-navigational-load");
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
