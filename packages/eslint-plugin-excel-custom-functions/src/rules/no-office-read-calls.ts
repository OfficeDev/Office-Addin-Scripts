import { TSESTree, ESLintUtils, TSESLint, AST_NODE_TYPES } from "@typescript-eslint/experimental-utils";
import { isCallSignatureDeclaration, isIdentifier } from "typescript";
import { getCustomFunction, isOfficeObject, isOfficeFuncWriteOrRead, OfficeCalls } from './utils'

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

                    if (isOfficeFuncWriteOrRead(node, typeChecker, services) === OfficeCalls.READ) {
                        const customFunction = getCustomFunction(services, ruleContext);
    
                        if (customFunction) {
                            ruleContext.report({
                                messageId: "officeReadCall",
                                loc: node.loc,
                                node: node
                            });
                        }

                    }
                }
            },

            AssignmentExpression: function(node: TSESTree.AssignmentExpression) {
                if (isOfficeObject(node.right, typeChecker, services)) {
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

            VariableDeclarator: function(node: TSESTree.VariableDeclarator) {
                if (node.init && isOfficeObject(node.init, typeChecker, services)) {
                    
                    const customFunction = getCustomFunction(services, ruleContext);

                    if (customFunction) {
                        ruleContext.report({
                            messageId: "officeReadCall",
                            loc: node.loc,
                            node: node
                        });
                    }
                }
            }
        };
    }
})



