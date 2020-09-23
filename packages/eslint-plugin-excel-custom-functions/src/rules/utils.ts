import { parse as parsePath } from 'path';
import {
  AST_NODE_TYPES,
  ESLintUtils,
  TSESLint,
  TSESTree
} from '@typescript-eslint/experimental-utils';
import * as ts from 'typescript';
let version = "0.0.8";

export enum OfficeCalls {
  WRITE = "WRITE",
  READ = "READ",
  NOTOFFICE = "NOTOFFICE"
}

type RequiredParserServices = ReturnType<typeof ESLintUtils.getParserServices>;
export type Options = unknown[];
export type MessageIds = string;

export const REPO_URL = 'https://github.com/arttarawork/Office-Addin-Scripts/packages/eslint-plugin-office-custom-functions';


export const createRule = ESLintUtils.RuleCreator(name => {
  const ruleName = parsePath(name).name;

  return `${REPO_URL}/blob/v${version}/docs/rules/${ruleName}.md`;
});

// Code to determine if function has @customfunction tag

export function isCustomFunction(node: TSESTree.Node, services: RequiredParserServices): boolean {
  const functionStarts = getAllFunctionStarts(node, services);
  for (let i = 0; i < functionStarts.length; i ++) {
    if (getJsDocCustomFunction(ts.getJSDocTags(functionStarts[i]))) {
      return true;
    }
  }

  return false;
}

// Maps all nodes in nodeArray to each other within helperFuncToHelperFuncMap
// Basically turns all nodes in NodeArray to this: https://en.wikipedia.org/wiki/Complete_graph
export function superNodeMe(nodeArray: Array<ts.Node>, helperFuncToHelperFuncMap: Map<ts.Node, Set<ts.Node>>): void {
  nodeArray.forEach((node, index) => {
    let currentVal = helperFuncToHelperFuncMap.get(node);
    if (!currentVal) {
      currentVal = new Set<ts.Node>([]);
    }
    for (let i = 0; i < nodeArray.length; i++) {
      if (i != index) {
        currentVal.add(nodeArray[i]);
      }
    }
    helperFuncToHelperFuncMap.set(
      node,
      currentVal
    );
  })
}

// Walks up the parent chain fron the current node to return all function declaration-like nodes that contain that initial node
export function getAllFunctionStarts(node: TSESTree.Node, services: RequiredParserServices): Array<ts.Node> {
  let tsNode: ts.Node = services.esTreeNodeToTSNodeMap.get(node);
  let outputArray: Array<ts.Node> = [];
  while (tsNode && tsNode.kind && !(tsNode.kind == ts.SyntaxKind.SourceFile)) {
    if (isFunctionDeclarationLike(tsNode)) {
      outputArray.push(tsNode);
    }
    tsNode = tsNode.parent;
  }
  return outputArray;
}

// Walks up the parent chain fron the current node to return the earliest function declaration-like node that contains that initial node
export function getStartOfFunction(node: TSESTree.Node, services: RequiredParserServices): ts.Node | undefined {
  let tsNode: ts.Node = services.esTreeNodeToTSNodeMap.get(node);
  while (tsNode && tsNode.kind && !(tsNode.kind == ts.SyntaxKind.SourceFile)) {
    if (isFunctionDeclarationLike(tsNode)) {
      return tsNode;
    }
    tsNode = tsNode.parent;
  }
  return undefined;
}

function isFunctionDeclarationLike(tsNode: ts.Node): boolean {
  return !!(tsNode && tsNode.kind 
    && (tsNode.kind == ts.SyntaxKind.CallSignature
    || tsNode.kind == ts.SyntaxKind.ConstructSignature
    || tsNode.kind == ts.SyntaxKind.MethodSignature
    || tsNode.kind == ts.SyntaxKind.IndexSignature
    || tsNode.kind == ts.SyntaxKind.FunctionType
    || tsNode.kind == ts.SyntaxKind.ConstructorType
    || tsNode.kind == ts.SyntaxKind.JSDocFunctionType
    || tsNode.kind == ts.SyntaxKind.FunctionDeclaration
    || tsNode.kind == ts.SyntaxKind.MethodDeclaration
    || tsNode.kind == ts.SyntaxKind.Constructor
    || tsNode.kind == ts.SyntaxKind.GetAccessor
    || tsNode.kind == ts.SyntaxKind.SetAccessor
    || tsNode.kind == ts.SyntaxKind.FunctionExpression
    || tsNode.kind == ts.SyntaxKind.ArrowFunction));
}

function getJsDocCustomFunction(tags: readonly ts.JSDocTag[]): { reason: string | (string & { __escapedIdentifier: void })} | undefined {
  for (const tag of tags) {
    if (tag.tagName.escapedText === 'customfunction') {
      return { reason: tag.tagName.escapedText || '' };
    }
  }
  return undefined;
}

// Code to determine if function is possibly write or read (Need new metadata pipeline to 100% determine, but this is a good heuristic for now):

export function isOfficeFuncWriteOrRead(node: TSESTree.CallExpression, typeChecker: ts.TypeChecker, services: RequiredParserServices): OfficeCalls | undefined {
  if (isOfficeObject(node.callee, typeChecker, services)) {
    let type = typeChecker.getTypeAtLocation(services.esTreeNodeToTSNodeMap.get(node.callee));
    let symbol = type.getSymbol();
    let symbolText = symbol ? typeChecker.symbolToString(symbol) : undefined;
    if (symbolText && 
      ( symbolText.toLowerCase().startsWith("set")
        || symbolText.toLowerCase().startsWith("add")
        || symbolText.toLowerCase().startsWith("clear")
        || symbolText.toLowerCase().startsWith("delete")
        || symbolText.toLowerCase().startsWith("remove")
        || symbolText.toLowerCase().startsWith("insert")
        || symbolText.toLowerCase().startsWith("copy")
        || symbolText.toLowerCase().startsWith("create")
      )
    ) {
      return OfficeCalls.WRITE;
    }

    let callSignatures = type.getCallSignatures();

    return callSignatures.some((callSignature) => {
      return (1 << 14 === ((1 << 14) & callSignature.getReturnType().flags.valueOf())); //bit-wise check to see if void is included in flags (See TypeFlags documentation in Typescript)
    }) ? OfficeCalls.WRITE : OfficeCalls.READ;
  }
  return undefined;
}

// Code to check if node is office object below:

export function isOfficeObject(node: TSESTree.Node, typeChecker: ts.TypeChecker, services: RequiredParserServices): boolean {
  let earlierMember: boolean = false;
  if (node.type == AST_NODE_TYPES.MemberExpression) {
    earlierMember = isOfficeObject(node.object, typeChecker, services);
  }
  if (earlierMember) {
    return true;
  }
  const officeDeclarations = getFunctionDeclarations(node, typeChecker, services);
  return officeDeclarations ? officeDeclarations.some(isParentNodeOfficeNamespace) : false;
}

function isParentNodeOfficeNamespace(node: ts.Node, index: number, decArray: ts.Declaration[]): boolean {
  const nodeText = node.getText();
  if (nodeText.startsWith("declare namespace Office")
    || nodeText.startsWith("declare namespace OfficeCore")
    || nodeText.startsWith("declare namespace Excel")) {
    return true;
  } else {
    return node.parent ? isParentNodeOfficeNamespace(node.parent, index, decArray) : false;
  }
}

// Code to get the function declaration some node is held within
// This will have to accomodate multiple func declarations
export function getFunctionDeclarations(node: TSESTree.Node, typeChecker: ts.TypeChecker, services: RequiredParserServices): ts.Declaration[] | undefined {
  if (node.type == AST_NODE_TYPES.CallExpression) {
    node = (<TSESTree.CallExpression>node).callee;
  }
  let tsNode = services.esTreeNodeToTSNodeMap.get(node);
  let type = typeChecker.getTypeAtLocation(tsNode);
  let symbol = type.getSymbol();
  if (!symbol) {
    symbol = typeChecker.getSymbolAtLocation(tsNode);
  }
  return symbol ? symbol.declarations : undefined;
}

export function isHelperFunc(node: TSESTree.Node, typeChecker: ts.TypeChecker, services: RequiredParserServices): boolean {
  const functionDeclarations = getFunctionDeclarations(node, typeChecker, services);
  let output = functionDeclarations ? functionDeclarations.some(
    (declaration) => {
      let sourceFile = declaration.getSourceFile();
      return !services.program.isSourceFileFromExternalLibrary(sourceFile);
    }
  ) : false;

  return output;
}

// Given a node and the helper function -> parent function map, return a set containing that original node and all parent funcs
export function bubbleUpNewCallingFuncs(node: ts.Node, helperFuncToHelperFuncMap: Map<ts.Node, Set<ts.Node>>): Set<ts.Node> {
  let outputSet = new Set<ts.Node>([node]);
  let examiningSet = helperFuncToHelperFuncMap.get(node);
  if (examiningSet) {
    examiningSet.forEach((nodeToExamine) => {
      let addingSet = helperFuncToHelperFuncMap.get(nodeToExamine);
      if (addingSet) {
        addingSet.forEach((nodeToAdd) => {
          if (!outputSet.has(nodeToAdd)) {
            examiningSet?.add(nodeToAdd);
          }
        });
      }
      outputSet.add(nodeToExamine)
      examiningSet?.delete(nodeToExamine);
    })
  }
  return outputSet;
}

// If a function (or any parent functions that use the original function as a helper function) has queued reports, report them all
export function reportIfCalledFromCustomFunction(nodeToBubbleUpFrom: ts.Node,
  ruleContext: TSESLint.RuleContext<MessageIds, Options>, 
  helperFuncToHelperFuncMap: Map<ts.Node, Set<ts.Node>>, 
  helperFuncToMentionsMap: Map<ts.Node, Array<{messageId: MessageIds, loc: TSESTree.SourceLocation, node: TSESTree.Node}>>,
  officeCallingFuncs?: Set<ts.Node>): void {
  bubbleUpNewCallingFuncs(nodeToBubbleUpFrom, helperFuncToHelperFuncMap).forEach((bubbledUp) => {
    if (officeCallingFuncs) {
      officeCallingFuncs.add(bubbledUp);
    }
    helperFuncToMentionsMap.get(bubbledUp)?.forEach((mention) => {
        ruleContext.report(mention);
    });
    helperFuncToMentionsMap.delete(bubbledUp);
  });
}

export function callExpressionAnalysis(node: TSESTree.CallExpression, 
  services: RequiredParserServices, 
  typeChecker: ts.TypeChecker, 
  ruleContext: TSESLint.RuleContext<MessageIds, Options>,
  officeCallingFuncs: Set<ts.Node>, 
  helperFuncToMentionsMap: Map<ts.Node, Array<{messageId: MessageIds, loc: TSESTree.SourceLocation, node: TSESTree.Node}>>, 
  helperFuncToHelperFuncMap: Map<ts.Node, Set<ts.Node>>,
  isCheckingForWrite: boolean = false): void {
  if (isOfficeObject(node, typeChecker, services)) {
    if (isOfficeFuncWriteOrRead(node, typeChecker, services) === (isCheckingForWrite ? OfficeCalls.WRITE : OfficeCalls.READ)) {
      if (isCustomFunction(node, services)) {
        
        // Reporting cases where office calls are made in custom functions
        ruleContext.report({
          messageId: isCheckingForWrite ? "officeWriteCall" : "officeReadCall",
          loc: node.loc,
          node: node
        });
      }

      // Adds the function this node is part of to a set of Office-calling functions
      // Releases all queued reports for that function and any parent functions that use it

      const functionStart = getStartOfFunction(node, services);
      if (functionStart) {
        reportIfCalledFromCustomFunction(functionStart, 
          ruleContext, 
          helperFuncToHelperFuncMap, 
          helperFuncToMentionsMap, 
          officeCallingFuncs
        );
      }
    }
  } else if (isHelperFunc(node, typeChecker, services)) {
    const functionDeclarations = getFunctionDeclarations(node, typeChecker, services);

    if (functionDeclarations && functionDeclarations.length > 0) {
      superNodeMe(functionDeclarations, helperFuncToHelperFuncMap);

      if(isCustomFunction(node, services)) {
        if (functionDeclarations.some((declaration) => {
          return officeCallingFuncs.has(declaration);
        })) {

          // If this is in a custom function and if this call expression is noted as having an office call in it, report
          ruleContext.report({
            messageId: isCheckingForWrite ? "officeWriteCall" : "officeReadCall",
            loc: node.loc,
            node: node
          });
        } else {

          // Otherwise, keep note of the location with a queued report
          // Should later analysis reveal that the call expression has an office call in it, we have the report ready
          helperFuncToMentionsMap.set(functionDeclarations[0], 
            (helperFuncToMentionsMap.get(functionDeclarations[0]) || []).concat({
              messageId: isCheckingForWrite ? "officeWriteCall" : "officeReadCall",
              loc: node.loc,
              node: node
            })
          );
        }
      }

      const functionStart = getStartOfFunction(node, services);
      
      if (functionStart) {

        // add an inner function -> parent function mapping for later tracking
        helperFuncToHelperFuncMap.set(
          functionDeclarations[0], 
          (helperFuncToHelperFuncMap.get(functionDeclarations[0]) || new Set<ts.Node>([])).add(functionStart)
        );

        // Releases all queued reports for that function and any parent functions that use it
        reportIfCalledFromCustomFunction(functionStart, 
          ruleContext, 
          helperFuncToHelperFuncMap, 
          helperFuncToMentionsMap
        );
      }
    }
  }
}

export function assignmentExpressionAnalysis(node: TSESTree.AssignmentExpression, ruleContext: TSESLint.RuleContext<MessageIds, Options>, services: RequiredParserServices, typeChecker: ts.TypeChecker, isCheckingForWrite: boolean = false): void {
  if (isOfficeObject(isCheckingForWrite ? node.left : node.right, typeChecker, services) && isCustomFunction(node, services)) {
    ruleContext.report({
      messageId: isCheckingForWrite ? "officeWriteCall" : "officeReadCall",
      loc: node.loc,
      node: node
    });
  }
}

export function variableDeclaratorAnalysis(node: TSESTree.VariableDeclarator, ruleContext: TSESLint.RuleContext<MessageIds, Options>, services: RequiredParserServices, typeChecker: ts.TypeChecker): void {
  if (node.init && isOfficeObject(node.init, typeChecker, services) && isCustomFunction(node, services)) {
    ruleContext.report({
      messageId: "officeReadCall",
      loc: node.loc,
      node: node
    });
  }
}
 