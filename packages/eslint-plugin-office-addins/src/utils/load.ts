import { AST_NODE_TYPES, TSESTree } from "@typescript-eslint/utils";
import { findCallExpression } from "./utils";

export function isLoadFunction(node: TSESTree.MemberExpression): boolean {
  const methodCall = findCallExpression(node);
  return methodCall !== undefined && isLoadCall(methodCall);
}

export function isLoadCall(node: TSESTree.CallExpression): boolean {
  return (
    node &&
    node.callee.type === AST_NODE_TYPES.MemberExpression &&
    node.callee.property.type === AST_NODE_TYPES.Identifier &&
    node.callee.property.name === "load"
  );
}

export function isLoadReference(
  node: TSESTree.Identifier | TSESTree.JSXIdentifier,
) {
  return (
    node.parent &&
    node.parent.type === TSESTree.AST_NODE_TYPES.MemberExpression &&
    isLoadFunction(node.parent)
  );
}

export function isContextLoadArgumentReference(
  node: TSESTree.Identifier | TSESTree.JSXIdentifier,
) {
  return (
    node.parent?.type === AST_NODE_TYPES.CallExpression &&
    node.parent.callee.type === AST_NODE_TYPES.MemberExpression &&
    node.parent.callee.object.type === AST_NODE_TYPES.Identifier &&
    node.parent.callee.object.name === "context" &&
    node.parent.callee.property.type === AST_NODE_TYPES.Identifier &&
    node.parent.callee.property.name === "load"
  );
}

function parseObjectExpressionProperty(
  objectExpression: TSESTree.ObjectExpression,
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
            propertyName + "/" + composedProperty,
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
  const methodCall = findCallExpression(node);

  if (methodCall && isLoadCall(methodCall)) {
    const argument = methodCall.arguments[0];
    if (!argument) {
      return [];
    }

    return parsePropertiesArgument(argument);
  }
  throw new Error("error in parseLoadArgument function.");
}

export function parsePropertiesArgument(
  argument: TSESTree.CallExpressionArgument,
): string[] {
  let properties: string[] = [];
  if (argument.type === AST_NODE_TYPES.ArrayExpression) {
    argument.elements.forEach((element) => {
      if (element != null && element.type === TSESTree.AST_NODE_TYPES.Literal) {
        properties = properties.concat(
          parseLoadStringArgument(element.value as string),
        );
      }
    });
  } else if (argument.type === TSESTree.AST_NODE_TYPES.Literal) {
    properties = parseLoadStringArgument(argument.value as string);
  } else if (argument.type === TSESTree.AST_NODE_TYPES.ObjectExpression) {
    properties = parseObjectExpressionProperty(argument);
  }

  return properties;
}
