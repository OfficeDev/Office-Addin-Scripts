import {
  AST_NODE_TYPES,
  TSESTree,
} from "@typescript-eslint/experimental-utils";
import { findTopLevelExpression } from "./utils";

export function isLoadFunction(node: TSESTree.MemberExpression): boolean {
  node = findTopLevelExpression(node);

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

export function getLoadArgument(
  node: TSESTree.MemberExpression
): string[] | undefined {
  node = findTopLevelExpression(node);

  if (
    isLoadFunction(node) &&
    node.parent?.type === TSESTree.AST_NODE_TYPES.CallExpression
  ) {
    if (
      node.parent.arguments.length === 0 ||
      (node.parent.arguments[0].type === TSESTree.AST_NODE_TYPES.Literal &&
        (node.parent.arguments[0].value as string) === "")
    ) {
      return undefined;
    }

    let properties: string[] = [];
    node.parent.arguments.forEach(
      (propertiesSeparatedByComma: TSESTree.CallExpressionArgument) => {
        if (
          propertiesSeparatedByComma.type === TSESTree.AST_NODE_TYPES.Literal
        ) {
          (propertiesSeparatedByComma.value as string)
            .replace(/\s/g, "")
            .split(",")
            .forEach((property: string) => {
              properties.push(property);
            });
        } else if (
          propertiesSeparatedByComma.type ===
          TSESTree.AST_NODE_TYPES.ObjectExpression
        ) {
          properties.push(
            composeObjectExpressionPropertyIntoString(
              propertiesSeparatedByComma
            )
          );
        }
      }
    );

    return properties;
  }
  throw new Error("error in getLoadArgument function.");
}
