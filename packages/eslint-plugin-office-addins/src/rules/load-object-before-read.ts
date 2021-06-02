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
      loadBeforeRead: "An explicit load call on '{{name}}' needs to be made before reading a proxy object",
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
    const SENTINEL_TYPE = /^(?:(?:Function|Class)(?:Declaration|Expression)|ArrowFunctionExpression|CatchClause|ImportDeclaration|ExportNamedDeclaration)$/u;
    const FOR_IN_OF_TYPE = /^For(?:In|Of)Statement$/u;
    
    /**
     * Checks whether or not a given location is inside of the range of a given node.
     * @param {ASTNode} node An node to check.
     * @param {number} location A location to check.
     * @returns {boolean} `true` if the location is inside of the range of the node.
     */
     function isInRange(node: TSESTree.Node, location: number): boolean {
      return node && node.range[0] <= location && location <= node.range[1];
    }
    
    /**
     * Checks whether or not a given reference is inside of the initializers of a given variable.
     *
     * This returns `true` in the following cases:
     *
     *     var a = a
     *     var [a = a] = list
     *     var {a = a} = obj
     *     for (var a in a) {}
     *     for (var a of a) {}
     * @param {Variable} variable A variable to check.
     * @param {Reference} reference A reference to check.
     * @returns {boolean} `true` if the reference is inside of the initializers.
     */
     function isInInitializer(variable: Variable, reference: Reference): boolean {
      if (variable.scope !== reference.from) {
          return false;
      }
    
      let node: TSESTree.Node | undefined = variable.identifiers[0].parent;
      const location: number = reference.identifier.range[1];
    
      while (node) {
          if (node.type === "VariableDeclarator") {
              if (node.init !== undefined && isInRange(node.init as TSESTree.Node, location)) {
                  return true;
              }
              /*if (FOR_IN_OF_TYPE.test(node.parent?.parent?.type as string) &&
                  isInRange(node.parent?.parent?.right, location)
              ) {
                  return true;
              }*/
              break;
          } else if (node.type === "AssignmentPattern") {
              if (isInRange(node.right, location)) {
                  return true;
              }
          } else if (SENTINEL_TYPE.test(node.type)) {
              break;
          }
    
          node = node.parent;
      }
    
      return false;
    }    

    /**
     * Finds if valueRead was loaded beforeHand
     * @param variable 
     * @returns
    */
    function isLoaded(referenceNode: Reference, valueRead: string): boolean {
      const variable = referenceNode.resolved;
      //console.log("Logging all the references:");
      variable?.references.forEach((reference: Reference) => {
        //console.log(reference.identifier);

        if(reference.identifier.parent?.type === "MemberExpression"
          && (reference.identifier.parent.property as TSESTree.Identifier).name === "load"
          && reference.identifier.parent.parent?.type === "CallExpression"
          && (reference.identifier.parent.parent.arguments[0] as TSESTree.Literal).value === valueRead) {

            if(reference.identifier.range[1] < referenceNode.identifier.range[1]) {
              return true;
            }
          }
      });

      return false;
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
        //console.log("Reference = ");
        //console.log(reference);
        if(variable?.name === "selectedRange" && !reference.init) {
          console.log(variable?.name);
          /*console.log(reference.init);
          console.log(variable?.identifiers.length);
          console.log(variable?.identifiers[0].range[1]);
          console.log(reference.identifier.range[1]);
          console.log(isInInitializer(variable as Variable, reference));*/
          console.log(isLoaded(reference, ""));
        }
        if (reference.init // ok
            || !variable // ok
            // || variable.identifiers.length === 0
            || (variable.identifiers[0].range[1] < reference.identifier.range[1]
            && !isInInitializer(variable, reference))
          ) {
            return;
        }
        // Add an If here to check if load is before reference.identifier.range[1]
        // Reports.
        context.report({
            node: reference.identifier,
            messageId: "loadBeforeRead",
            data: {name: reference.identifier.name}
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


/*
    return {
      ":matches(VariableDeclarator[init.callee.property.name = 'getSelectedRange'], VariableDeclarator[init.callee.property.name = 'getItem'], VariableDeclarator[init.callee.property.name = 'getRange'])"(
          node: TSESTree.VariableDeclarator
      ) {
        const variableName: string = (node.id as TSESTree.Identifier).name;

        context.report({
          node: node,
          messageId: "loadBeforeRead",
        });
      },
    };
*/

/*
Location the load function:
CallExpression[callee.property.name='load'][arguments.Literal.value = variableName]
*/
