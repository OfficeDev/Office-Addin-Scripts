import { TSESTree } from "@typescript-eslint/typescript-estree";
import {
  Reference,
  Scope,
  Variable,
} from "@typescript-eslint/experimental-utils/dist/ts-eslint-scope";
import { BooleanArraySupportOption } from "prettier";
import { stringify } from "querystring";

export = {
  name: "load-object-before-read",
  meta: {
    type: <"problem" | "suggestion" | "layout">"problem",
    messages: {
      loadBeforeRead: "An explicit load call on '{{name}}' for '{{loadValue}}' needs to be made before reading a proxy object",
    },
    docs: {
      description: 
        "Before you can read the properties of a proxy object, you must explicitly load the properties",
      category: <
        "Best Practices" | "Stylistic Issues" | "Variables" | "Possible Errors"
      >"Possible Errors",
      recommended: <false | "error" | "warn">false,
      url: 
        "https://docs.microsoft.com/en-us/office/dev/add-ins/develop/application-specific-api-model?view=powerpoint-js-1.1#load",
    },
    schema: [],
  },
  create: function (context: any) {
    function getPropertyThatHadToBeLoaded(node: TSESTree.Node): string | undefined {
      if(node.parent?.type === TSESTree.AST_NODE_TYPES.MemberExpression
        && node.parent?.property.type === TSESTree.AST_NODE_TYPES.Identifier) {
        return node.parent.property.name;
      }
      return undefined;
    }

    function isLoadFunction(node: TSESTree.Node): boolean {
      return (node.parent?.type === TSESTree.AST_NODE_TYPES.MemberExpression
        && node.parent.property.type === TSESTree.AST_NODE_TYPES.Identifier
        && node.parent.property.name === "load");
    }

    function isGetFunction(node: TSESTree.Identifier): boolean {
      const functionName = node.name;
      return (functionName === "getSelectedRange"
        || functionName === "getItem" 
        || functionName === "getRange");
    }

    function isAGetFunction2(node: TSESTree.Node): boolean {
      let getFunctionFound = false;
      if(node.parent?.type === TSESTree.AST_NODE_TYPES.VariableDeclarator
        && node.parent.init?.type === TSESTree.AST_NODE_TYPES.CallExpression
        && node.parent.init.callee.type === TSESTree.AST_NODE_TYPES.MemberExpression
        && node.parent.init.callee.property.type === TSESTree.AST_NODE_TYPES.Identifier) {
          if(isGetFunction(node.parent.init.callee.property)) {
            getFunctionFound = true;
          }
      }
      return getFunctionFound;
    }

    function getLoadedPropertyName(node: TSESTree.Node): string {
      if(node.parent?.parent?.type === TSESTree.AST_NODE_TYPES.CallExpression
        && node.parent.parent.arguments[0].type === TSESTree.AST_NODE_TYPES.Literal) {
          return node.parent.parent.arguments[0].value as string;
        }
      return "error in getLoadedPropertyName";
    }

    function findLoadBeforeRead(scope: Scope) {
      scope.variables.forEach((variable: Variable) => {
        let loadLocation: Map <string, number> = new Map<string, number>();
        let getFound: boolean = false;
        variable.references.forEach((reference: Reference) => {
          if (reference.init
            || !variable) {
            return;
          }

          if (isLoadFunction(reference.identifier)) {
            loadLocation.set(getLoadedPropertyName(reference.identifier), reference.identifier.range[1]);
            return;
          }

          if (isAGetFunction2(reference.identifier)) {
            getFound = true;
            return;
          }

          // If reference came after load 
          const propertyName: string | undefined = getPropertyThatHadToBeLoaded(reference.identifier);
          if (!propertyName) {
            return;
          }
          if (loadLocation.has(propertyName)
            && (reference.identifier.range[1] > (loadLocation.get(propertyName) ?? 0)) ) {
              return;
          }

          context.report({
            node: reference.identifier,
            messageId: "loadBeforeRead",
            data: {name: reference.identifier.name, loadValue: propertyName}
          });
        });
      });
      scope.childScopes.forEach(findLoadBeforeRead);
    }

    return {
      Program() {
        findLoadBeforeRead(context.getScope());
      }
    }
  },
};
