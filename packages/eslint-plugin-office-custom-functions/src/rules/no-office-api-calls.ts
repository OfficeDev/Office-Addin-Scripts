import { TSESTree, ESLintUtils, TSESLint, AST_NODE_TYPES } from "@typescript-eslint/experimental-utils";
import { isCallSignatureDeclaration, isIdentifier } from "typescript";
import { isOfficeBoilerplate, getCustomFunction2, isOfficeObject } from './utils'

/**
 * @fileoverview Prevents office api calls
 * @author Artur Tarasenko (artarase)
 */
"use strict";

const createRule = ESLintUtils.RuleCreator(
  () => 'https://github.com/OfficeDev/Office-Addin-Scripts',
);


//------------------------------------------------------------------------------
// Rule Definition
//------------------------------------------------------------------------------

let excelRunToContextMap = new Map<TSESTree.Node, TSESTree.Identifier>();
let contextToExcelRunMap = new Map<TSESTree.Node, TSESTree.Node>();

function isInExcelRun(node: TSESTree.Node): TSESTree.Node | undefined {
    if (excelRunToContextMap.has(node)) {
        return node;
    } else {
        return node.parent ? isInExcelRun(node.parent) : undefined;
    }
}

export type Options = unknown[];
export type MessageIds = 'contextSync';

export default createRule<Options, MessageIds>({
    name: 'no-office-api-calls',

    meta: {
        docs: {
            description: "Prevents office api calls",
            category: "Best Practices",
            recommended: "error",
            requiresTypeChecking: true
        },
        type: "problem",
        messages: {
            contextSync: "No context.sync() calls within Custom Functions"
        },
        schema: []
    },

    defaultOptions: [],
        
    create(ruleContext) {
        const services = ESLintUtils.getParserServices(ruleContext);

        return {
            CallExpression: function(node: TSESTree.CallExpression) {
                if(isOfficeBoilerplate(node)) {
                    if(node.arguments[0].type == "FunctionExpression"
                        && node.arguments[0].params.length > 0
                        && node.arguments[0].params[0].type == "Identifier") {
                        excelRunToContextMap.set(node, node.arguments[0].params[0]);
                        contextToExcelRunMap.set(node.arguments[0].params[0], node);
                    }
                }
            },

            Identifier: function(node: TSESTree.Identifier) {
                let excelRunNode = isInExcelRun(node);
                let originalContext: TSESTree.Identifier | undefined;

                if (!!excelRunNode && excelRunToContextMap.has(excelRunNode)) {

                    originalContext = excelRunToContextMap.get(excelRunNode);
                    if (originalContext?.name == node.name) {
                        contextToExcelRunMap.set(node, excelRunNode);
                        if (node.parent
                            && node.parent.type == "MemberExpression"
                            && node.parent.property.type == "Identifier"
                            && node.parent.property.name == "sync") {
                                const customFunction = getCustomFunction2(services, ruleContext);

                                if (customFunction) {
                                    ruleContext.report({
                                        messageId: "contextSync",
                                        loc: node.parent.loc,
                                        node: node.parent
                                    });
                                }
                        }
                    }
                }
            },

            AssignmentExpression: function(node: TSESTree.AssignmentExpression) {
                if (isInExcelRun(node) || isOfficeObject(node)) {
                    const customFunction = getCustomFunction2(services, ruleContext);

                    if (customFunction) {
                        ruleContext.report({
                            messageId: "contextSync",
                            loc: node.loc,
                            node: node
                        });
                    }
                }
            }
        };
    }
})



