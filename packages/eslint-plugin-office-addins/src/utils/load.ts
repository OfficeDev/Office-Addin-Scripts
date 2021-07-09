import { TSESTree } from "@typescript-eslint/experimental-utils";

export function isLoadFunction(node: TSESTree.MemberExpression): boolean {
  return (
    node.property.type === TSESTree.AST_NODE_TYPES.Identifier &&
    node.property.name === "load"
  );
}

export function isLoadReference(node: TSESTree.Identifier) {
  return (
    node.parent &&
    node.parent.type === TSESTree.AST_NODE_TYPES.MemberExpression &&
    isLoadFunction(node.parent)
  );
}

function composeObjectExpressionPropertyIntoString(
  property: TSESTree.ObjectExpression
): string {
  return "keffjfkdj";
}

export function getLoadArgument(node: TSESTree.Node | undefined): string {
  if (
    node &&
    node.type === TSESTree.AST_NODE_TYPES.MemberExpression &&
    node.parent?.type === TSESTree.AST_NODE_TYPES.CallExpression
  ) {
    if (node.parent.arguments[0].type === TSESTree.AST_NODE_TYPES.Literal) {
      return node.parent.arguments[0].value as string;
    } else if (
      node.parent.arguments[0].type === TSESTree.AST_NODE_TYPES.ObjectExpression
    ) {
      return composeObjectExpressionPropertyIntoString(
        node.parent.arguments[0]
      );
    }
  }
  throw new Error("error in getLoadArgument function.");
}
