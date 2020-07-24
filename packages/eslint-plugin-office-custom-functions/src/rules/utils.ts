import { parse as parsePath } from 'path';
import {
  AST_NODE_TYPES,
  ESLintUtils,
  TSESLint,
  TSESTree,
} from '@typescript-eslint/experimental-utils';
import { isReassignmentTarget, getJsDoc } from 'tsutils';
import * as ts from 'typescript';
import { Scope } from '@typescript-eslint/experimental-utils/dist/ts-eslint';
import * as metadata from "../data/metadata.json";
let version = "0.0.8";

const REPO_URL = 'https://github.com/arttarawork/Office-Addin-Scripts';

export const createRule = ESLintUtils.RuleCreator(name => {
  const ruleName = parsePath(name).name;

  return `${REPO_URL}/packages/eslint-plugin-office-custom-functions/blob/v${version}/docs/rules/${ruleName}.md`;
});

export const isOfficeBoilerplate = (node: TSESTree.CallExpression) => {
    return node.type == "CallExpression"
        && !!node.callee 
        && node.callee.type == "MemberExpression"
        && node.callee.property.type == "Identifier"
        && node.callee.property.name == "run"
        && node.callee.object.type == "Identifier"
        && isOfficeNamespace(node.callee.object)
}

export const isOfficeNamespace = (node: TSESTree.Identifier) => {
  return ( node.name == "Excel"
    || node.name == "Word"
    || node.name == "Powerpoint"
  )
}

export const isOfficeMemberOrCallExpression = (node: TSESTree.MemberExpression | TSESTree.CallExpression): boolean => {
  if (node.type == "CallExpression" && node.callee && node.callee.type == "MemberExpression") {
    return isOfficeMemberOrCallExpression(node.callee);
  } else if (node.type == "MemberExpression") {
    if (node.object.type == "MemberExpression") {
      return isOfficeMemberOrCallExpression(node.object);
    } else if (node.object.type == "Identifier") {
      return isOfficeNamespace(node.object);
    }
  }
  return false;
}

export const isDescendantOf = (descendantNode: TSESTree.Node, ancestorNode: TSESTree.Node) : boolean => {
    if (descendantNode.parent === ancestorNode) {
        return true;
    } else {
        return descendantNode.parent ? isDescendantOf(descendantNode.parent, ancestorNode) : false;
    }
}

//Requires more work
export const isCustomFunction = (node: TSESTree.Node, context: TSESLint.RuleContext<any,any>) : boolean => {
    return !!context.getSourceCode().getJSDocComment(node);
}

type RequiredParserServices = ReturnType<typeof ESLintUtils.getParserServices>;
export type Options = unknown[];
export type MessageIds = string;

export function getCustomFunction(
  services: RequiredParserServices,
  context: TSESLint.RuleContext<MessageIds, Options>,
) {
  const functionStarts = getFunctionStarts(context);
  for (let i = 0; i < functionStarts.length; i ++) {
    const tsNode = getTsNode(functionStarts[i], services);
    if (tsNode) {
      const JSDocTags = ts.getJSDocTags(tsNode);
      const customFunction = getJsDocCustomFunction(JSDocTags);
      if (customFunction) {
        return customFunction;
      }
    }
  }

  return undefined;
}

function getFunctionStarts(
  context: TSESLint.RuleContext<MessageIds, Options>,
): Array<TSESTree.Node> {
  let outputArray: Array<TSESTree.Node> = [];
  const ancestors = context.getAncestors();
  for (let i = 0; i < ancestors.length; i++) {
    if (ancestors[i].type == "ExportNamedDeclaration"
      && (i + 1) < ancestors.length 
      && ancestors[i + 1].type == "FunctionDeclaration") {
        outputArray.push(ancestors[i]);
        i++;
    } else if (ancestors[i].type == "FunctionDeclaration") {
      outputArray.push(ancestors[i]);
    }
  }
  return outputArray;
}

function getTsNode(
  node: TSESTree.Node,
  services: RequiredParserServices,
) {
  const tsNode = services.esTreeNodeToTSNodeMap.get(
    node as TSESTree.Node,
  ) as ts.Node;
  return tsNode

}

function getJsDocCustomFunction(tags: readonly ts.JSDocTag[]) {
  for (const tag of tags) {
    if (tag.tagName.escapedText === 'customfunction') {
      return { reason: tag.tagName.escapedText || '' };
    }
  }
  return undefined;
}

export function isOfficeObject(node: TSESTree.Node | null, map?: Map<TSESTree.Node, any>) {
  if (!!node) {
    if (node.type == "CallExpression" || node.type == "MemberExpression") {
      const isOfficeNamespace = isOfficeMemberOrCallExpression(node);
      if (isOfficeNamespace) {
        return true;
      }
    } if (map && map.has(node)) {
      return true;
    }
  }
  return false;
}

enum OfficeCalls {
  WRITE = "WRITE",
  READ = "READ"
}

export function isOfficeFunctionCall(node: TSESTree.CallExpression): OfficeCalls | undefined {
  return undefined
}

// export function compareAndAdd(node: TSESTree.Node, map: Map<TSESTree.Node, any>) {
//   if ()
// }

export function checkOfficeCall(node: TSESTree.CallExpression): OfficeCalls | undefined {
  if (node.callee.type == "MemberExpression") {
    if (isOfficeObject(node.callee.object)) {
      if (node.callee.property.type == "Identifier") {
        if (node.callee.property.name.startsWith("set")) {
          return OfficeCalls.WRITE;
        } else if (node.callee.property.name.startsWith("get")) {
          return OfficeCalls.READ;
        }
      }
    }
  }

  return undefined;
}


// let officeObjectTracker: Array<Map<TSESTree.Node, any>> = [];

// export function addToOfficeDictionary (
//   node: TSESTree.Node, 
//   officeObjectTracker: Map<Scope.Scope, Map<TSESTree.Node, any>>, 
//   ruleContext: TSESLint.RuleContext<string, unknown[]>
// ): void {
//   const scope = ruleContext.getScope();
//   if (officeObjectTracker && officeObjectTracker.has(scope)) {
//     let innerMap = officeObjectTracker.get(scope);
//     if (innerMap) {
//       officeObjectTracker.set(scope, innerMap.set(node, "OFFICE"))
//     }
//   } else {
//     officeObjectTracker.set(scope, new Map<TSESTree.Node, any>().set(node, "OFFICE"));
//   }

// }

export function addToOfficeDictionary (
  node: TSESTree.Node, 
  officeObjectTracker: Map<Scope.Scope, Map<string, string>>, 
  ruleContext: TSESLint.RuleContext<string, unknown[]>,
  services: RequiredParserServices = ESLintUtils.getParserServices(ruleContext),
): void {
  const scope = ruleContext.getScope();
  if (officeObjectTracker && officeObjectTracker.has(scope)) {
    let innerMap = officeObjectTracker.get(scope);
    if (innerMap) {
      officeObjectTracker.set(scope, innerMap.set(getTsNode(node, services).getText(), "OFFICE"));
    }
  } else {
    officeObjectTracker.set(scope, new Map<string, string>().set(getTsNode(node, services).getText(), "OFFICE"));
  }
}

export function hasInOfficeDictionary (
  node: TSESTree.Node, 
  officeObjectTracker: Map<Scope.Scope, Map<string, string>>, 
  ruleContext: TSESLint.RuleContext<string, unknown[]>,
  services: RequiredParserServices = ESLintUtils.getParserServices(ruleContext),
): boolean {
  let currentScope: Scope.Scope | null = ruleContext.getScope();
  while (currentScope) {
    if (officeObjectTracker.get(currentScope)?.has(getTsNode(node, services).getText())) {
      return true;
    }
    currentScope = currentScope.upper;
  }
  return false;
}

export function getFromOfficeDictionary (
  node: TSESTree.Node, 
  officeObjectTracker: Map<Scope.Scope, Map<string, string>>, 
  ruleContext: TSESLint.RuleContext<string, unknown[]>,
  services: RequiredParserServices = ESLintUtils.getParserServices(ruleContext),
): any {
  let currentScope: Scope.Scope | null = ruleContext.getScope();
  while (currentScope) {
    let possibleOutput = officeObjectTracker.get(currentScope)?.get(getTsNode(node, services).getText());
    if (possibleOutput) {
      return possibleOutput;
    }
    currentScope = currentScope.upper;
  }
  return undefined;
}