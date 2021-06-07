import { TSESTree, AST_NODE_TYPES } from "@typescript-eslint/typescript-estree";
import {
  Reference,
  Scope,
  Variable,
} from "@typescript-eslint/experimental-utils/dist/ts-eslint-scope";
import {
  RuleFix,
  RuleFixer,
} from "@typescript-eslint/experimental-utils/dist/ts-eslint";

export = {
  name: "test-for-null-using-isNullObject",
  meta: {
    type: <"problem" | "suggestion" | "layout">"problem",
    messages: {
      useIsNullObject: "Test '{{name}}' for null using isNullObject.",
    },
    docs: {
      description: "",
      category: <
        "Best Practices" | "Stylistic Issues" | "Variables" | "Possible Errors"
      >"Possible Errors",
      recommended: <false | "error" | "warn">false,
      url: "https://docs.microsoft.com/en-us/office/dev/add-ins/develop/application-specific-api-model?view=powerpoint-js-1.1#ornullobject-methods-and-properties",
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

    function isNullObjectCall(node: TSESTree.Node): boolean {
      if (
        node.type === AST_NODE_TYPES.CallExpression &&
        node.callee.type === AST_NODE_TYPES.MemberExpression &&
        node.callee.property.type === AST_NODE_TYPES.Identifier &&
        node.callee.property.name.endsWith("OrNullObject")
      ) {
        return true;
      }
      return false;
    }

    function isNullObjectNode(node: TSESTree.Node | undefined): boolean {
      if (
        node &&
        ((node.type === AST_NODE_TYPES.VariableDeclarator &&
          node.init &&
          isNullObjectCall(node.init) &&
          node.id.type === AST_NODE_TYPES.Identifier) ||
          (node.type === AST_NODE_TYPES.AssignmentExpression &&
            isNullObjectCall(node.right) &&
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
        let nullTests: Reference[] = [];

        for (let ref = 0; ref < references.length; ref++) {
          const reference = references[ref];

          if (isNullObjectNode(reference.identifier.parent)) {
            nullObjectCall = true;
          }

          if (isInNullTest(reference.identifier)) {
            nullTests.push(reference);
          }
        }

        if (nullObjectCall === true && nullTests.length > 0) {
          nullTests.forEach((reference) => {
            context.report({
              node: reference.identifier,
              messageId: "useIsNullObject",
              data: { name: reference.identifier.name },
              fix: function (fixer: RuleFixer) {
                var ruleFix: RuleFix;
                if (
                  isInBinaryNullTest(reference.identifier) &&
                  reference.identifier.parent
                ) {
                  let newTest = reference.identifier.name + ".isNullObject";
                  ruleFix = fixer.replaceText(
                    reference.identifier.parent,
                    newTest
                  );
                } else {
                  ruleFix = fixer.insertTextAfter(
                    reference.identifier,
                    ".isNullObject"
                  );
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
      "Program:exit"(programNode: TSESTree.Node) {
        findNullObjectNullTests(context.getScope());
      },
    };
  },
};
