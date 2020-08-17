import { TSESTree, ESLintUtils, TSESLint, AST_NODE_TYPES } from "@typescript-eslint/experimental-utils";
//import * as sm from "@typescript-eslint/scope-manager";
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

                if(isOfficeObject(node, typeChecker, services)) {

                    if (isOfficeFuncWriteOrRead(node, typeChecker, services) === OfficeCalls.WRITE) {
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
            },

            AssignmentExpression: function(node: TSESTree.AssignmentExpression) {
                if (isOfficeObject(node.left, typeChecker, services)) {
                    
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



