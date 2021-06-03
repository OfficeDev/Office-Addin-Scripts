import { TSESTree } from "@typescript-eslint/typescript-estree";
import {
  Reference,
  Variable,
} from "@typescript-eslint/experimental-utils/dist/ts-eslint-scope";

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
          if (FOR_IN_OF_TYPE.test(node.parent?.parent?.type as string) &&
              isInRange(node.parent?.parent?.right, location)
          ) {
              return true;
          }
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



export = {
  name: "load-object-before-read",
  meta: {
    type: <"problem" | "suggestion" | "layout">"problem",
    messages: {
      loadBeforeRead: "An explicit load call needs to be made before reading a proxy object",
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
          if (reference.init ||
              !variable ||
              variable.identifiers.length === 0 ||
              (variable.identifiers[0].range[1] < reference.identifier.range[1] && !isInInitializer(variable, reference))
            ) {
              return;
          }

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
