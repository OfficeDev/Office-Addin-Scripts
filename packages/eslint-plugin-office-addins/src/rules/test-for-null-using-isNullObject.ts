import {
  TSESTree,
  AST_NODE_TYPES,
} from "@typescript-eslint/experimental-utils";
import {
  Reference,
  Scope,
  Variable,
} from "@typescript-eslint/experimental-utils/dist/ts-eslint-scope";
import {
  RuleFix,
  RuleFixer,
} from "@typescript-eslint/experimental-utils/dist/ts-eslint";
import { isGetOrNullObjectFunction } from "../utils/getFunction";

export = {
  name: "test-for-null-using-isNullObject",
  meta: {
    type: <"problem" | "suggestion" | "layout">"problem",
    messages: {
      useIsNullObject: "Test the isNullObject property of '{{name}}'.",
    },
    docs: {
      description:
        "Do not test the truthiness of an object returned by an OrNullObject method or property. Test it's isNullObject property.",
      category: <
        "Best Practices" | "Stylistic Issues" | "Variables" | "Possible Errors"
      >"Possible Errors",
      recommended: <false | "error" | "warn">false,
      url: "https://docs.microsoft.com/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties",
    },
    schema: [],
    fixable: <"code" | "whitespace">"code",
  },
  create: function (context: any) {
    function isConditionalTestExpression(node: TSESTree.Identifier): boolean {
      return (
        node.parent != undefined &&
        (node.parent.type === AST_NODE_TYPES.IfStatement ||
          node.parent.type === AST_NODE_TYPES.WhileStatement ||
          node.parent.type === AST_NODE_TYPES.DoWhileStatement ||
          node.parent.type === AST_NODE_TYPES.ForStatement ||
          node.parent.type === AST_NODE_TYPES.ConditionalExpression) &&
        node === node.parent.test
      );
    }

    function isInUnaryNullTest(node: TSESTree.Identifier): boolean {
      return (
        node.parent != undefined &&
        node.parent.type === AST_NODE_TYPES.UnaryExpression &&
        node.parent.operator === "!" &&
        node.parent.argument === node
      );
    }

    function isInBinaryNullTest(node: TSESTree.Identifier): boolean {
      return (
        node.parent != undefined &&
        node.parent.type === AST_NODE_TYPES.BinaryExpression &&
        ((node.parent.left === node &&
          node.parent.right.type === AST_NODE_TYPES.Literal &&
          node.parent.right.raw === "null") ||
          (node.parent.right === node &&
            node.parent.left.type === AST_NODE_TYPES.Literal &&
            node.parent.left.raw === "null"))
      );
    }

    function isInNullTest(node: TSESTree.Identifier): boolean {
      return (
        isConditionalTestExpression(node) ||
        node.parent?.type === AST_NODE_TYPES.LogicalExpression ||
        isInUnaryNullTest(node) ||
        isInBinaryNullTest(node)
      );
    }

    function isNullObjectNode(node: TSESTree.Node | undefined): boolean {
      if (
        node &&
        ((node.type === AST_NODE_TYPES.VariableDeclarator &&
          node.init &&
          isGetOrNullObjectFunction(node.init) &&
          node.id.type === AST_NODE_TYPES.Identifier) ||
          (node.type === AST_NODE_TYPES.AssignmentExpression &&
            isGetOrNullObjectFunction(node.right) &&
            node.left.type === AST_NODE_TYPES.Identifier))
      ) {
        return true;
      }
      return false;
    }

    function findNullObjectNullTests(scope: Scope): void {
      const variables = scope.variables;
      const childScopes = scope.childScopes;

      for (let i = 0; i < variables.length; i++) {
        const variable: Variable = variables[i];
        const references: Reference[] = variable.references;
        let nullObjectCall: boolean = false;
        const nullTests: TSESTree.Identifier[] = [];

        for (let ref = 0; ref < references.length; ref++) {
          const identifier: TSESTree.Identifier = references[ref].identifier;

          if (isNullObjectNode(identifier.parent)) {
            nullObjectCall = true;
          }

          if (isInNullTest(identifier)) {
            nullTests.push(identifier);
          }
        }

        if (nullObjectCall === true && nullTests.length > 0) {
          nullTests.forEach((identifier) => {
            context.report({
              node: identifier,
              messageId: "useIsNullObject",
              data: { name: identifier.name },
              fix: function (fixer: RuleFixer) {
                var ruleFix: RuleFix;
                if (isInBinaryNullTest(identifier) && identifier.parent) {
                  let newTest = identifier.name + ".isNullObject";
                  ruleFix = fixer.replaceText(identifier.parent, newTest);
                } else {
                  ruleFix = fixer.insertTextAfter(identifier, ".isNullObject");
                }
                return ruleFix;
              },
            });
          });
        }
      }

      for (let i = 0; i < childScopes.length; ++i) {
        findNullObjectNullTests(childScopes[i]);
      }
    }

    return {
      "Program:exit"(
        programNode: TSESTree.Node /* eslint-disable-line no-unused-vars */
      ) {
        findNullObjectNullTests(context.getScope());
      },
    };
  },
};
