import {
  AST_NODE_TYPES,
  TSESTree,
} from "@typescript-eslint/experimental-utils";

function findMemberExpresionBeforeLoad(node: TSESTree.Node | undefined): TSESTree.MemberExpression {
  if (node &&
    node.type === TSESTree.AST_NODE_TYPES.MemberExpression) {
      while (node.parent && 
        node.parent.type === AST_NODE_TYPES.MemberExpression) {
        node = node.parent;
      }
    } else {
      throw "Error in findMemberExpresionBeforeLoad";
    }

  return node;
}

export function isLoadFunction(node: TSESTree.MemberExpression): boolean {
  node = findMemberExpresionBeforeLoad(node);

  return (
    node.parent?.type === AST_NODE_TYPES.CallExpression &&
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
  objectExpression: TSESTree.ObjectExpression
): string {
  let composedProperty: string = "";
  objectExpression.properties.forEach((property) => {
    if (property.type === AST_NODE_TYPES.Property) {
      if (property.key.type === AST_NODE_TYPES.Identifier) {
        composedProperty += property.key.name;
      }
      if (property.value.type === AST_NODE_TYPES.ObjectExpression) {
        composedProperty +=
          "/" + composeObjectExpressionPropertyIntoString(property.value);
      }
    }
  });

  return composedProperty;
}

export function getLoadArgument(node: TSESTree.Node | undefined): string {
  node = findMemberExpresionBeforeLoad(node);

  if (
    isLoadFunction(node) &&
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
