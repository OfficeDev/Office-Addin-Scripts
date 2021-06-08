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

    function callsAGetAPIFunction(node: TSESTree.Identifier): boolean {
      const functionName = node.name;
      return (functionName === "getSelectedRange"
        || functionName === "getItem" 
        || functionName === "getRange");
    }

    function isAGetVariableDeclaration(node: TSESTree.Node): boolean {
      if(node.parent?.type === TSESTree.AST_NODE_TYPES.VariableDeclarator
        && node.parent.init?.type === TSESTree.AST_NODE_TYPES.CallExpression
        && node.parent.init.callee.type === TSESTree.AST_NODE_TYPES.MemberExpression
        && node.parent.init.callee.property.type === TSESTree.AST_NODE_TYPES.Identifier) {
          if(callsAGetAPIFunction(node.parent.init.callee.property)) {
            return true;
          }
      }
      return false;
    }

    function isAGetAssignmentExpression(node: TSESTree.Node): boolean {
      if(node.parent?.type === TSESTree.AST_NODE_TYPES.AssignmentExpression
        && node.parent.right.type === TSESTree.AST_NODE_TYPES.CallExpression
        && node.parent.right.callee.type === TSESTree.AST_NODE_TYPES.MemberExpression
        && node.parent.right.callee.property.type === TSESTree.AST_NODE_TYPES.Identifier) {
          if(callsAGetAPIFunction(node.parent.right.callee.property)) {
            return true;
          }
      }
      return false;
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

          const node: TSESTree.Node = reference.identifier;
          if (isLoadFunction(node)) {
            loadLocation.set(getLoadedPropertyName(node), node.range[1]);
            return;
          }

          if (isAGetVariableDeclaration(node)
              || isAGetAssignmentExpression(node)) {
            getFound = true;
            return;
          }

          // If reference came after load 
          const propertyName: string | undefined = getPropertyThatHadToBeLoaded(node);
          if (!propertyName) {
            return;
          }

          if (loadLocation.has(propertyName)
            && (node.range[1] > (loadLocation.get(propertyName) ?? 0))) {
              return;
          }

          context.report({
            node: node,
            messageId: "loadBeforeRead",
            data: {name: node.name, loadValue: propertyName}
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
