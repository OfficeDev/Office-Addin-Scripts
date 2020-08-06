import { TSESTree, ESLintUtils, TSESLint, AST_NODE_TYPES } from "@typescript-eslint/experimental-utils";
import { isCallSignatureDeclaration, isIdentifier } from "typescript";
import { isOfficeBoilerplate, getCustomFunction, isOfficeObject, isOfficeNamespace, isOfficeMemberOrCallExpression } from './utils'

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
export type MessageIds = 'officeReadCall';

export default createRule<Options, MessageIds>({
    name: 'no-office-read-calls',

    meta: {
        docs: {
            description: "Prevents office api calls",
            category: "Best Practices",
            recommended: "warn",
            requiresTypeChecking: true
        },
        type: "problem",
        messages: {
            officeReadCall: "No Office API read calls within Custom Functions"
        },
        schema: []
    },

    defaultOptions: [],
        
    create(ruleContext) {
        const services = ESLintUtils.getParserServices(ruleContext);

        const typeChecker = services.program.getTypeChecker();

        return {
            CallExpression: function(node: TSESTree.CallExpression) {
                if (isOfficeObject(node, typeChecker, services)) {
                    const customFunction = getCustomFunction(services, ruleContext);

                    if (customFunction) {
                        ruleContext.report({
                            messageId: "officeReadCall",
                            loc: node.loc,
                            node: node
                        });
                    }
                }
            },

            // Identifier: function(node: TSESTree.Identifier) {
            //     let excelRunNode = isInExcelRun(node);
            //     let originalContext: TSESTree.Identifier | undefined;

            //     if (!!excelRunNode && excelRunToContextMap.has(excelRunNode)) {

            //         originalContext = excelRunToContextMap.get(excelRunNode);
            //         if (originalContext?.name == node.name) {
            //             contextToExcelRunMap.set(node, excelRunNode);
            //         }
            //     }
            // },
            MemberExpression: function(node: TSESTree.MemberExpression) {
                if (isOfficeObject(node, typeChecker, services)) {
                    const customFunction = getCustomFunction(services, ruleContext);

                    if (customFunction) {
                        ruleContext.report({
                            messageId: "officeReadCall",
                            loc: node.loc,
                            node: node
                        });
                    }
                }
            },

            // AssignmentExpression: function(node: TSESTree.AssignmentExpression) {
            //     if (node.right.type == "ArrayExpression") {
            //         for (let i = 0; i < node.right.elements.length; i++) {
            //             if(isOfficeMemberOrCallExpression(node.right.elements))
            //         }
            //     }
            //     else if (isOfficeMemberOrCallExpression(node.right))
            //     if (isInExcelRun(node) || isOfficeObject(node)) {
            //         const customFunction = getCustomFunction(services, ruleContext);

            //         if (customFunction) {
            //             ruleContext.report({
            //                 messageId: "officeReadCall",
            //                 loc: node.loc,
            //                 node: node
            //             });
            //         }
            //     }
            // },

            // VariableDeclaration: function(node: TSESTree.VariableDeclaration) {
            //     for (let i = 0; i < node.declarations.length; i++) {
            //         if (node.declarations[i].)
            //     }
            // }
        };
    }
})



