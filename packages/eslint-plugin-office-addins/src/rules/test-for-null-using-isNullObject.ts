import {
  ESLintUtils,
  TSESTree,
  AST_NODE_TYPES,
} from "@typescript-eslint/utils";
import { Reference, Scope, Variable } from "@typescript-eslint/scope-manager";
import { RuleFix, RuleFixer } from "@typescript-eslint/utils/ts-eslint";
import { isGetOrNullObjectFunction } from "../utils/getFunction";

export default ESLintUtils.RuleCreator(
  () =>
    "https://docs.microsoft.com/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties",
)({
  name: "test-for-null-using-isNullObject",
  meta: {
    type: "problem",
    messages: {
      useIsNullObject: "Test the isNullObject property of '{{name}}'.",
    },
    docs: {
      description:
        "Do not test the truthiness of an object returned by an OrNullObject method or property. Test it's isNullObject property.",
    },
    schema: [],
    fixable: <"code" | "whitespace">"code",
  },
  create: function (context) {
    const sourceCode = context.sourceCode ?? context.getSourceCode();
    function isConditionalTestExpression(
      node: TSESTree.Identifier | TSESTree.JSXIdentifier,
    ): boolean {
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

    function isInUnaryNullTest(
      node: TSESTree.Identifier | TSESTree.JSXIdentifier,
    ): boolean {
      return (
        node.parent != undefined &&
        node.parent.type === AST_NODE_TYPES.UnaryExpression &&
        node.parent.operator === "!" &&
        node.parent.argument === node
      );
    }

    function isInBinaryNullTest(
      node: TSESTree.Identifier | TSESTree.JSXIdentifier,
    ): boolean {
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

    function isInNullTest(
      node: TSESTree.Identifier | TSESTree.JSXIdentifier,
    ): boolean {
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
        const nullTests: (TSESTree.Identifier | TSESTree.JSXIdentifier)[] = [];

        for (let ref = 0; ref < references.length; ref++) {
          const identifier: TSESTree.Identifier | TSESTree.JSXIdentifier =
            references[ref].identifier;

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
                let ruleFix: RuleFix;
                if (isInBinaryNullTest(identifier) && identifier.parent) {
                  const newTest = identifier.name + ".isNullObject";
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
      "Program:exit"(node) {
        const scope = sourceCode.getScope
          ? sourceCode.getScope(node)
          : context.getScope();
        findNullObjectNullTests(scope);
      },
    };
  },
  defaultOptions: [],
});
