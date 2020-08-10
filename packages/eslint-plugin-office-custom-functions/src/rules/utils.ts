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
const metadata: {[key: string]: {comment: Array<string>, attributes: any, properties: Array<any>, methods: Array<any>} } = require("../data/metadata.json");
let version = "0.0.8";
// import * as office from "@microsoft/office-js"

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

// export function isOfficeObject(node: TSESTree.Node | null, map?: Map<TSESTree.Node, any>) {
//   if (!!node) {
//     if (node.type == "CallExpression" || node.type == "MemberExpression") {
//       const isOfficeNamespace = isOfficeMemberOrCallExpression(node);
//       if (isOfficeNamespace) {
//         return true;
//       }
//     } if (map && map.has(node)) {
//       return true;
//     }
//   }
//   return false;
// }

export enum OfficeCalls {
  WRITE = "WRITE",
  READ = "READ"
}

export class OfficeFunctionReturns {
  public name: string;
  public returnType: string;
  public callType: OfficeCalls;

  constructor(
    funcName: string,
    funcReturnType: string,
    funcCallType: OfficeCalls
  ) {
    this.name = funcName;
    this.returnType = funcReturnType;
    this.callType = funcCallType;
  }
}

export function isOfficeFunctionCall(node: TSESTree.CallExpression): OfficeCalls | undefined {
  return undefined
}

// export function compareAndAdd(node: TSESTree.Node, map: Map<TSESTree.Node, any>) {
//   if ()
// }

// export function checkOfficeCall(node: TSESTree.CallExpression): OfficeCalls | undefined {
//   if (node.callee.type == "MemberExpression") {
//     if (isOfficeObject(node.callee.object)) {
//       if (node.callee.property.type == "Identifier") {
//         if (node.callee.property.name.startsWith("set")) {
//           return OfficeCalls.WRITE;
//         } else if (node.callee.property.name.startsWith("get")) {
//           return OfficeCalls.READ;
//         }
//       }
//     }
//   }

//   return undefined;
// }

export function getOfficeObject(type: string) : {comment: Array<string>, attributes: any, properties: Array<any>, methods: Array<any>} | undefined {
  return metadata[type];
}

function officePropertiesFunctionChecker(properties: any[], func: string, isItSetter: boolean): boolean {
  return properties.some(
    (property) =>
    {
      return property && property.name 
        && property.name == func.slice(3) 
        && (isItSetter ? property.set : property.get);
    }
  );
}

function officeMethodFunctionChecker(methods: any[], func: string): OfficeCalls | undefined {
  methods.forEach((method) => {
    if (method && method.name && method.name == func) {
      if (method.type && method.type == "void") {
        return OfficeCalls.WRITE;
      }
      if (<string>(method.name).startsWith("set")) {
        return OfficeCalls.WRITE;
      }
      return OfficeCalls.READ;
    }
  })
  return undefined;
}

export function isFuncInOfficeType(type: string, func: string): OfficeCalls | undefined {
  const typeObject = getOfficeObject(type);
  if (typeObject) {

    if (func.startsWith("set") && officePropertiesFunctionChecker(typeObject.properties, func, true)) {
      return OfficeCalls.WRITE;
    }
    //add, delete, clear

    const methodFunc = officeMethodFunctionChecker(typeObject.methods, func);

    if (methodFunc) {
      return methodFunc;
    }

    if (func.startsWith("get") && officePropertiesFunctionChecker(typeObject.properties, func, false)) {
      return OfficeCalls.READ;
    }
  }
  return undefined;
}

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

function isParentNodeOfficeNamespace(node: ts.Node, index: number, decArray: ts.Declaration[]): boolean {
  const nodeText = node.getText();
  if (
    nodeText.startsWith("declare namespace Office")
    || nodeText.startsWith("declare namespace OfficeCore")
    || nodeText.startsWith("declare namespace Excel")
    || nodeText.startsWith("declare namespace Word")
    || nodeText.startsWith("declare namespace OneNote")
    || nodeText.startsWith("declare namespace Visio")
    || nodeText.startsWith("declare namespace PowerPoint")
    || nodeText.startsWith("declare namespace Project")
  ) {
    return true;
  } else {
    return node.parent ? isParentNodeOfficeNamespace(node.parent, index, decArray) : false;
  }
}

export function isOfficeObject(node: TSESTree.Node, typeChecker: ts.TypeChecker, services: RequiredParserServices): boolean {
  let type = typeChecker.getTypeAtLocation(services.esTreeNodeToTSNodeMap.get(node));
  let symbol = type.getSymbol();
  return (symbol && symbol.declarations) ? symbol.declarations.some(isParentNodeOfficeNamespace) : false;
}

export function getOfficeFuncReturnType(type: string, func: string): string | undefined {
  const typeObject = getOfficeObject(type);
  if (typeObject) {
    if (typeObject.methods) {
      typeObject.methods.forEach((method) => {
        if (method && method.name && method.name == func && method.type) {
          return <string>(method.type)
        }
      });
    }
  }
  //add prop method check
  return undefined;
}