import {
  AST_NODE_TYPES,
  TSESTree,
} from "@typescript-eslint/utils";
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

function parseObjectExpressionProperty(
  objectExpression: TSESTree.ObjectExpression
): string[] {
  let composedProperties: string[] = [];

  objectExpression.properties.forEach((property) => {
    if (
      property.type === AST_NODE_TYPES.Property &&
      property.key.type === AST_NODE_TYPES.Identifier
    ) {
      let propertyName: string = property.key.name;

      if (property.value.type === AST_NODE_TYPES.ObjectExpression) {
        const composedProperty = parseObjectExpressionProperty(property.value);
        if (composedProperty.length !== 0) {
          composedProperties = composedProperties.concat(
            propertyName + "/" + composedProperty
          );
        }
      } else if (
        property.value.type === AST_NODE_TYPES.Literal &&
        property.value.value // Checking if the value assigined to the property is true
      ) {
        composedProperties = composedProperties.concat(propertyName);
      }
    }
  });

  return composedProperties;
}

function parseLoadStringArgument(argument: string): string[] {
  let properties: string[] = [];
  argument
    .replace(/\s/g, "")
    .split(",")
    .forEach((property: string) => {
      properties.push(property);
    });

  return properties;
}

export function parseLoadArguments(node: TSESTree.MemberExpression): string[] {
  node = findTopLevelExpression(node);

  if (
    isLoadFunction(node) &&
    node.parent?.type === TSESTree.AST_NODE_TYPES.CallExpression
  ) {
    const argument = node.parent.arguments[0];
    if (!argument) {
      return [];
    }

    let properties: string[] = [];
    if (argument.type === AST_NODE_TYPES.ArrayExpression) {
      argument.elements.forEach((element) => {
        if (element.type === TSESTree.AST_NODE_TYPES.Literal) {
          properties = properties.concat(
            parseLoadStringArgument(element.value as string)
          );
        }
      });
    } else if (argument.type === TSESTree.AST_NODE_TYPES.Literal) {
      properties = properties.concat(
        parseLoadStringArgument(argument.value as string)
      );
    } else if (argument.type === TSESTree.AST_NODE_TYPES.ObjectExpression) {
      properties = properties.concat(parseObjectExpressionProperty(argument));
    }

    return properties;
  }
  throw new Error("error in getLoadArgument function.");
}
