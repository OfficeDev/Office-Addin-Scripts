import { TSESTree } from "@typescript-eslint/typescript-estree";
import {
  Reference,
  Variable,
} from "@typescript-eslint/experimental-utils/dist/ts-eslint-scope";

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
    function getValueThatHadToBeLoaded(referenceNode: Reference): string | undefined {
      if(referenceNode.identifier.parent?.type === TSESTree.AST_NODE_TYPES.MemberExpression
        && referenceNode.identifier.parent?.property.type === TSESTree.AST_NODE_TYPES.Identifier) {
        return referenceNode.identifier.parent.property.name;
      }
      return undefined;
    }

    function isLoadFunction(reference: Reference): boolean {
      if(reference.identifier.parent?.type === TSESTree.AST_NODE_TYPES.MemberExpression
        && reference.identifier.parent.property.type === TSESTree.AST_NODE_TYPES.Identifier
        && reference.identifier.parent.property.name === "load") {
        return true;
      }
      return false;
    }

    function wasCreatedByGetFunction(referenceNode: Reference): boolean {
      const variable = referenceNode.resolved;
      let getFunctionFound = false;
      variable?.references.forEach((reference: Reference) => {
          if(reference.identifier.parent?.type === TSESTree.AST_NODE_TYPES.VariableDeclarator
            && reference.identifier.parent.init?.type === TSESTree.AST_NODE_TYPES.CallExpression
            && reference.identifier.parent.init.callee.type === TSESTree.AST_NODE_TYPES.MemberExpression
            && reference.identifier.parent.init.callee.property.type === TSESTree.AST_NODE_TYPES.Identifier) {
              const functionName = reference.identifier.parent.init.callee.property.name;
              if(functionName === "getSelectedRange"
                || functionName === "getItem" 
                || functionName === "getRange") {
                getFunctionFound = true;
              }
          }
        });
      return getFunctionFound;
    }

    function isLoaded(referenceNode: Reference): boolean {
      const variable = referenceNode.resolved;
      let loadFound = false;
      const valueRead = getValueThatHadToBeLoaded(referenceNode);

      variable?.references.forEach((reference: Reference) => {
        if(reference === referenceNode) return;
        if(reference.identifier.parent?.parent?.type === TSESTree.AST_NODE_TYPES.CallExpression
          && reference.identifier.parent.parent.arguments[0].type === TSESTree.AST_NODE_TYPES.Literal
          && reference.identifier.parent.parent.arguments[0].value === valueRead
          && reference.identifier.range[1] < referenceNode.identifier.range[1]) {
          loadFound = true;
        }
      });

      return loadFound;
    }

    function findVariablesInScope(scope: any) {
      scope.references.forEach((reference: Reference) => {
        const variable = reference.resolved;

        if (reference.init
            || !variable
            || !wasCreatedByGetFunction(reference)
            || isLoadFunction(reference)
            || isLoaded(reference)){
            return;
        }
        context.report({
            node: reference.identifier,
            messageId: "loadBeforeRead",
            data: {name: reference.identifier.name, loadValue: getValueThatHadToBeLoaded(reference)}
        });
      });
      scope.childScopes.forEach(findVariablesInScope);
    }

    return {
      Program() {
        findVariablesInScope(context.getScope());
      }
    }
  },
};
