import { TSESTree } from "@typescript-eslint/experimental-utils";
import {
  Reference,
  Scope,
  Variable,
} from "@typescript-eslint/experimental-utils/dist/ts-eslint-scope";
import { parseLoadArguments, isLoadFunction } from "../utils/load";
import { findPropertiesRead } from "../utils/utils";
import { isGetFunction } from "../utils/getFunction";

export = {
  name: "load-object-before-read",
  meta: {
    type: <"problem" | "suggestion" | "layout">"problem",
    messages: {
      loadBeforeRead:
        "An explicit load call on '{{name}}' for property '{{loadValue}}' needs to be made before the property can be read.",
    },
    docs: {
      description:
        "Before you can read the properties of a proxy object, you must explicitly load the properties.",
      category: <
        "Best Practices" | "Stylistic Issues" | "Variables" | "Possible Errors"
      >"Possible Errors",
      recommended: <false | "error" | "warn">false,
      url: "https://docs.microsoft.com/office/dev/add-ins/develop/application-specific-api-model#load",
    },
    schema: [],
  },
  create: function (context: any) {
    function isInsideWriteStatement(node: TSESTree.Node): boolean {
      while (node.parent) {
        node = node.parent;
        if (node.type === TSESTree.AST_NODE_TYPES.AssignmentExpression)
          return true;
      }
      return false;
    }

    function findLoadBeforeRead(scope: Scope) {
      scope.variables.forEach((variable: Variable) => {
        let loadLocation: Map<string, number> = new Map<string, number>();
        let getFound: boolean = false;
        let loadStar: boolean = false;

        variable.references.forEach((reference: Reference) => {
          const node: TSESTree.Node = reference.identifier;

          if (
            node.parent?.type === TSESTree.AST_NODE_TYPES.VariableDeclarator
          ) {
            getFound = false; // In case of reassignment

            if (node.parent.init && isGetFunction(node.parent.init)) {
              getFound = true;
              return;
            }
          }

          if (
            node.parent?.type === TSESTree.AST_NODE_TYPES.AssignmentExpression
          ) {
            getFound = false; // In case of reassignment

            if (isGetFunction(node.parent.right)) {
              getFound = true;
              return;
            }
          }

          if (!getFound) {
            // If reference was not related to a previous get
            return;
          }

          if (node.parent?.type === TSESTree.AST_NODE_TYPES.MemberExpression) {
            if (isLoadFunction(node.parent)) {
              // In case it is a load function
              const propertyNames: string[] = parseLoadArguments(node.parent);
              propertyNames.forEach((propertyName: string) => {
                if (propertyName === "*") {
                  loadStar = true;
                } else {
                  loadLocation.set(propertyName, node.range[1]);
                }
              });
              return;
            }
          }

          const propertyName: string | undefined = findPropertiesRead(
            node.parent
          );
          if (!propertyName) {
            // There is no property
            return;
          }

          if (
            loadLocation.has(propertyName) && // If reference came after load, return
            node.range[1] > (loadLocation.get(propertyName) ?? 0)
          ) {
            return;
          }

          if (loadStar) {
            return;
          }

          if (isInsideWriteStatement(node)) {
            // Return in case it a write, ie, not read statment
            return;
          }

          context.report({
            node: node,
            messageId: "loadBeforeRead",
            data: { name: node.name, loadValue: propertyName },
          });
        });
      });
      scope.childScopes.forEach(findLoadBeforeRead);
    }

    return {
      Program() {
        findLoadBeforeRead(context.getScope());
      },
    };
  },
};
