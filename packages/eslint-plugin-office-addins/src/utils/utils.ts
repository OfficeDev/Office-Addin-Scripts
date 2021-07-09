import { TSESTree } from "@typescript-eslint/experimental-utils";
import {
  Reference,
  Scope,
  Variable,
} from "@typescript-eslint/experimental-utils/dist/ts-eslint-scope";
import { isGetFunction } from "./getFunction";

export function isLoadFunction(node: TSESTree.MemberExpression): boolean {
  return (
    node.property.type === TSESTree.AST_NODE_TYPES.Identifier &&
    node.property.name === "load"
  );
}

export function isContextSyncIdentifier(node: TSESTree.Identifier): boolean {
  return (
    node.name === "context" &&
    node.parent?.type === TSESTree.AST_NODE_TYPES.MemberExpression &&
    node.parent?.parent?.type === TSESTree.AST_NODE_TYPES.CallExpression &&
    node.parent?.property.type === TSESTree.AST_NODE_TYPES.Identifier &&
    node.parent?.property.name === "sync"
  );
}

export type OfficeApiReference = {
  operation: "Read" | "Load" | "Write" | "Sync";
  reference: Reference;
};

export function isLoadReference(node: TSESTree.Identifier) {
  return (
    node.parent &&
    node.parent.type === TSESTree.AST_NODE_TYPES.MemberExpression &&
    isLoadFunction(node.parent)
  );
}

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
      apiReferences.push({ operation: "Write", reference: reference });
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

export function findPropertiesRead(node: TSESTree.Node | undefined): string {
  let propertyName = ""; // Will be a string combined with '/' for the case of navigation properties
  while (node) {
    if (
      node.type === TSESTree.AST_NODE_TYPES.MemberExpression &&
      node.property.type === TSESTree.AST_NODE_TYPES.Identifier
    ) {
      propertyName += node.property.name + "/";
    }
    node = node.parent;
  }
  return propertyName.slice(0, -1);
}

export function getLiteralArgumentName(
  node: TSESTree.Node | undefined
): string {
  if (
    node &&
    node.type === TSESTree.AST_NODE_TYPES.MemberExpression &&
    node.parent?.type === TSESTree.AST_NODE_TYPES.CallExpression &&
    node.parent.arguments[0].type === TSESTree.AST_NODE_TYPES.Literal
  ) {
    return node.parent.arguments[0].value as string;
  }
  throw new Error("error in getLiteralArgumentName function.");
}
