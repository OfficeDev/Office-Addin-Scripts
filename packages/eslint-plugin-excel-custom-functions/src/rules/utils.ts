import { parse as parsePath } from 'path';
import {
  AST_NODE_TYPES,
  ESLintUtils,
  TSESLint,
  TSESTree,
} from '@typescript-eslint/experimental-utils';
import * as ts from 'typescript';
// import * as metadata from "../data/metadata.json";
// const metadata: {[key: string]: {comment: Array<string>, attributes: any, properties: Array<any>, methods: Array<any>} } = require("../data/metadata.json");
let version = "0.0.8";
// import * as office from "@microsoft/office-js"

export enum OfficeCalls {
  WRITE = "WRITE",
  READ = "READ",
  NOTOFFICE = "NOTOFFICE"
}

type RequiredParserServices = ReturnType<typeof ESLintUtils.getParserServices>;
export type Options = unknown[];
export type MessageIds = string;

const REPO_URL = 'https://github.com/arttarawork/Office-Addin-Scripts';


export const createRule = ESLintUtils.RuleCreator(name => {
  const ruleName = parsePath(name).name;

  return `${REPO_URL}/packages/eslint-plugin-office-custom-functions/blob/v${version}/docs/rules/${ruleName}.md`;
});

//Code to determine if function has @customfunction tag

export function isCustomFunction(services: RequiredParserServices, 
  context: TSESLint.RuleContext<MessageIds, Options>): boolean {
  const functionStarts = getFunctionStarts(context);
  for (let i = 0; i < functionStarts.length; i ++) {
    if (services.esTreeNodeToTSNodeMap.get(functionStarts[i])) {
      const JSDocTags = ts.getJSDocTags(services.esTreeNodeToTSNodeMap.get(functionStarts[i]));
      if (getJsDocCustomFunction(JSDocTags)) {
        return true;
      }
    }
  }

  return false;
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

export function superNodeMe(
  nodeArray: Array<ts.Node>,
  helperFuncToHelperFuncMap: Map<ts.Node, Set<ts.Node>>
): void {
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

export function getFunctionStarts2(node: TSESTree.Node, services: RequiredParserServices) {
  let tsNode: ts.Node = services.esTreeNodeToTSNodeMap.get(node);
  let outputArray: Array<ts.Node> = [];
  while (tsNode && tsNode.kind && !(tsNode.kind == ts.SyntaxKind.SourceFile)) {
    if (tsNode.kind == ts.SyntaxKind.CallSignature
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
      || tsNode.kind == ts.SyntaxKind.ArrowFunction) {
        outputArray.push(tsNode);
      }
    tsNode = tsNode.parent;
  }
  return outputArray;
}

export function ancestorTextChain(node: TSESTree.Node, services: RequiredParserServices): Array<string> {
  let tsNode: ts.Node = services.esTreeNodeToTSNodeMap.get(node);

  let textArray: Array<string> = new Array<string>();

  while (tsNode) {
    textArray.push(tsNode.getText());
    tsNode = tsNode.parent;
  }
  
  return textArray;
}

export function ancestorChain(node: TSESTree.Node, services: RequiredParserServices): Array<ts.Node> {
  let tsNode: ts.Node = services.esTreeNodeToTSNodeMap.get(node);

  let textArray: Array<ts.Node> = new Array<ts.Node>();

  while (tsNode) {
    textArray.push(tsNode);
    tsNode = tsNode.parent;
  }
  
  return textArray;
}

function getJsDocCustomFunction(tags: readonly ts.JSDocTag[]) {
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
export function getFunctionDeclarations(node: TSESTree.Node, typeChecker: ts.TypeChecker, services: RequiredParserServices) {
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