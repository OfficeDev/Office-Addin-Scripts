import { TSESTree } from "@typescript-eslint/experimental-utils";
import * as getJson from "./data/getFunctions.json";

const getFunctions: Set<string> = new Set<string>(getJson.getFunctions);
const getOrNullObjectFunctions: Set<string> = new Set<string>(
  getJson.getOrNullObjectFunctions
);

export function isGetFunction(node: TSESTree.Node): boolean {
  return (
    node.type == TSESTree.AST_NODE_TYPES.CallExpression &&
    node.callee.type === TSESTree.AST_NODE_TYPES.MemberExpression &&
    node.callee.property.type === TSESTree.AST_NODE_TYPES.Identifier &&
    (getFunctions.has(node.callee.property.name) ||
      getOrNullObjectFunctions.has(node.callee.property.name))
  );
}

export function isGetOrNullObjectFunction(node: TSESTree.Node): boolean {
  return (
    node.type == TSESTree.AST_NODE_TYPES.CallExpression &&
    node.callee.type === TSESTree.AST_NODE_TYPES.MemberExpression &&
    node.callee.property.type === TSESTree.AST_NODE_TYPES.Identifier &&
    getOrNullObjectFunctions.has(node.callee.property.name)
  );
}
