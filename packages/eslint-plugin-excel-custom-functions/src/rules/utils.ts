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
// import * as metadata from "../data/metadata.json";
// const metadata: {[key: string]: {comment: Array<string>, attributes: any, properties: Array<any>, methods: Array<any>} } = require("../data/metadata.json");
let version = "0.0.8";
// import * as office from "@microsoft/office-js"

export enum OfficeCalls {
  WRITE = "WRITE",
  READ = "READ"
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

export function getCustomFunction(
  services: RequiredParserServices,
  context: TSESLint.RuleContext<MessageIds, Options>,
) {
  const functionStarts = getFunctionStarts(context);
  for (let i = 0; i < functionStarts.length; i ++) {
    if (services.esTreeNodeToTSNodeMap.get(functionStarts[i])) {
      const JSDocTags = ts.getJSDocTags(services.esTreeNodeToTSNodeMap.get(functionStarts[i]));
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
  if (node.type == AST_NODE_TYPES.CallExpression) {
    node = (<TSESTree.CallExpression>node).callee;
  }
  let tsNode = services.esTreeNodeToTSNodeMap.get(node);
  let type = typeChecker.getTypeAtLocation(tsNode);
  let symbol = type.getSymbol();
  if (!symbol) {
    symbol = typeChecker.getSymbolAtLocation(tsNode);
  }
  return (symbol && symbol.declarations) ? symbol.declarations.some(isParentNodeOfficeNamespace) : false;
}

function isParentNodeOfficeNamespace(node: ts.Node, index: number, decArray: ts.Declaration[]): boolean {
  const nodeText = node.getText();
  if (
    nodeText.startsWith("declare namespace Office")
    || nodeText.startsWith("declare namespace OfficeCore")
    || nodeText.startsWith("declare namespace Excel")
  ) {
    return true;
  } else {
    return node.parent ? isParentNodeOfficeNamespace(node.parent, index, decArray) : false;
  }
}

// Code to check if node is office object above ^