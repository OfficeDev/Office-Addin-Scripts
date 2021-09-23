import {
  AST_NODE_TYPES,
  TSESTree,
} from "@typescript-eslint/experimental-utils";
import {
  Reference,
  Scope,
  Variable,
} from "@typescript-eslint/experimental-utils/dist/ts-eslint-scope";
import { isGetFunction } from "./getFunction";
import { isLoadReference } from "./load";

export function isContextSyncIdentifier(node: TSESTree.Identifier): boolean {
  return (
    node.name === "context" &&
    node.parent?.type === AST_NODE_TYPES.MemberExpression &&
    node.parent?.parent?.type === AST_NODE_TYPES.CallExpression &&
    node.parent?.property.type === AST_NODE_TYPES.Identifier &&
    node.parent?.property.name === "sync"
  );
}

export function findTopLevelExpression(
  node: TSESTree.MemberExpression
): TSESTree.MemberExpression {
  while (node.parent && node.parent.type === AST_NODE_TYPES.MemberExpression) {
    node = node.parent;
  }

  return node;
}

export type OfficeApiReference = {
  operation: "Get" | "Load" | "Sync" | "Read";
  reference: Reference;
};

let proxyVariables: Set<Variable>;
let apiReferences: OfficeApiReference[];
export function findOfficeApiReferences(scope: Scope): OfficeApiReference[] {
  proxyVariables = new Set<Variable>();
  apiReferences = [];
  findOfficeApiReferencesInScope(scope);
  return apiReferences;
}

function findOfficeApiReferencesInScope(scope: Scope): void {
  scope.references.forEach((reference) => {
    if (
      reference.isWrite() &&
      reference.writeExpr &&
      isGetFunction(reference.writeExpr) &&
      reference.resolved
    ) {
      proxyVariables.add(reference.resolved);
      apiReferences.push({ operation: "Get", reference: reference });
    } else if (isContextSyncIdentifier(reference.identifier)) {
      apiReferences.push({ operation: "Sync", reference: reference });
    } else if (
      reference.isRead() &&
      reference.resolved &&
      proxyVariables.has(reference.resolved)
    ) {
      if (isLoadReference(reference.identifier)) {
        apiReferences.push({ operation: "Load", reference: reference });
      } else {
        apiReferences.push({ operation: "Read", reference: reference });
      }
    }
  });

  scope.childScopes.forEach(findOfficeApiReferencesInScope);
}

export function isProperty(node: TSESTree.Node): boolean {
  if (node.parent && node.parent.type === AST_NODE_TYPES.CallExpression) {
    const callExpression: TSESTree.Node = node.parent;

    let foundProperty = false;
    callExpression.arguments.forEach(
      (argument: TSESTree.CallExpressionArgument) => {
        if (argument === node) {
          foundProperty = true;
        }
      }
    );
    return foundProperty;
  }
  return true;
}

export function findPropertiesRead(node: TSESTree.Node | undefined): string {
  let propertyName = ""; // Will be a string combined with '/' for the case of navigation properties
  while (node) {
    if (
      node.type === AST_NODE_TYPES.MemberExpression &&
      node.property.type === AST_NODE_TYPES.Identifier &&
      isProperty(node)
    ) {
      propertyName += node.property.name + "/";
    }
    node = node.parent;
  }
  return propertyName.slice(0, -1);
}
