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
    function getLoadedValue(referenceNode: Reference): string {
      return ((referenceNode.identifier.parent as TSESTree.MemberExpression).property as TSESTree.Identifier).name;
    }

    function isLoaded(referenceNode: Reference): boolean {
      const variable = referenceNode.resolved;
      let loadFound = false;
      const valueRead = getLoadedValue(referenceNode);
      console.log("On isLoaded");
      variable?.references.forEach((reference: Reference) => {
        console.log("On a new reference");
        if(reference.identifier.parent?.type === "MemberExpression"
          && (reference.identifier.parent.property as TSESTree.Identifier).name === "load"
          && reference.identifier.parent.parent?.type === "CallExpression"
          && (reference.identifier.parent.parent.arguments[0] as TSESTree.Literal).value === valueRead
          && reference.identifier.range[1] < referenceNode.identifier.range[1]) {
          loadFound = true;
        }
        /*
        if(reference.identifier.parent?.type === "MemberExpression") {
          console.log("Inside first if");
          if((reference.identifier.parent.property as TSESTree.Identifier).name === "load") {
            console.log("Inside second if");
            if(reference.identifier.parent.parent?.type === "CallExpression") {
              console.log("Inside third if");
              if((reference.identifier.parent.parent.arguments[0] as TSESTree.Literal).value === valueRead) {
                console.log("Inside fourth if");
                if(reference.identifier.range[1] < referenceNode.identifier.range[1]) {
                  console.log("Inside fifth if");
                }
              }
            }
          }
        }*/
      });

      return loadFound;
    }

    /**
    * Finds and validates all variables in a given scope.
    * @param {Scope} scope The scope object.
    * @returns {void}
    * @private
    */
    function findVariablesInScope(scope: any) {
      scope.references.forEach((reference: Reference) => {
        const variable = reference.resolved;

        /*
          * Skips when the reference is:
          * - initialization's.
          * - referring to an undefined variable.
          * - referring to a global environment variable (there're no identifiers).
          * - located preceded by the variable (except in initializers).
          */

        /*if(variable?.name === "selectedRange" && !reference.init
          && ((reference.identifier.parent as any).property as TSESTree.Identifier).name !== "load") {
          console.log(isLoaded(reference));
        }*/
        if (reference.init
            || !variable
            || !isLoaded(reference)) {
            return;
        }
        // Reports.
        context.report({
            node: reference.identifier,
            messageId: "loadBeforeRead",
            data: {name: reference.identifier.name, loadValue: getLoadedValue(reference)}
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
