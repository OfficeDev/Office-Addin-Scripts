import { TSESTree, ESLintUtils, TSESLint, AST_NODE_TYPES } from "@typescript-eslint/experimental-utils";
//import * as sm from "@typescript-eslint/scope-manager";
import { isCallSignatureDeclaration, isIdentifier } from "typescript";
import { isOfficeBoilerplate, getCustomFunction, isOfficeObject, testboi, printboi } from './utils'

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
export type MessageIds = 'officeWriteCall';

export default createRule<Options, MessageIds>({
    name: 'no-office-write-calls',

    meta: {
        docs: {
            description: "Prevents office api calls",
            category: "Best Practices",
            recommended: "error",
            requiresTypeChecking: true
        },
        type: "problem",
        messages: {
            officeWriteCall: "No Office API write calls within Custom Functions"
        },
        schema: []
    },

    defaultOptions: [],
        
    create(ruleContext) {

        const services = ESLintUtils.getParserServices(ruleContext);

        const typeChecker = services.program.getTypeChecker();

        return {
            CallExpression: function(node: TSESTree.CallExpression) {

                if(isOfficeObject(node)) {
                    
                    const customFunction = getCustomFunction(services, ruleContext);

                    if (customFunction) {
                        ruleContext.report({
                            messageId: "officeWriteCall",
                            loc: node.loc,
                            node: node
                        });
                    }
                }
            },

            Identifier: function(node: TSESTree.Identifier) {
                printboi(node, typeChecker, services);


            },

            AssignmentExpression: function(node: TSESTree.AssignmentExpression) {
                if (isOfficeObject(node.right)) {
                    
                    const customFunction = getCustomFunction(services, ruleContext);

                    if (customFunction) {
                        ruleContext.report({
                            messageId: "officeWriteCall",
                            loc: node.loc,
                            node: node
                        });
                    }
                }
            },

            VariableDeclarator: function(node: TSESTree.VariableDeclarator) {
                if (isOfficeObject(node.init)) {
                    
                    const customFunction = getCustomFunction(services, ruleContext);

                    if (customFunction) {
                        ruleContext.report({
                            messageId: "officeWriteCall",
                            loc: node.loc,
                            node: node
                        });
                    }
                }
            }
        };
    }
})



