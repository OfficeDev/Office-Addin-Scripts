import { TSESTree, ESLintUtils } from "@typescript-eslint/experimental-utils";
import { isCustomFunction, isOfficeObject, isOfficeFuncWriteOrRead, OfficeCalls, getFunctionStarts2, ancestorChain, ancestorTextChain } from './utils'
import { RuleContext, RuleMetaDataDocs, RuleMetaData  } from '@typescript-eslint/experimental-utils/dist/ts-eslint';

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

type Options = unknown[];
type MessageIds = 'officeReadCall';

export = {
    name: 'no-office-read-calls',

    meta: {
        docs: {
            description: "Prevents office api calls",
            category: <RuleMetaDataDocs["category"]>"Best Practices",
            recommended: <RuleMetaDataDocs["recommended"]>"warn",
            requiresTypeChecking: true,
            url: 'https://github.com/OfficeDev/Office-Addin-Scripts'
        },
        type: <RuleMetaData<MessageIds>["type"]> "problem",
        messages: <Record<MessageIds, string>> {
            officeReadCall: "No Office API read calls within Custom Functions"
        },
        schema: []
    },

    create: function(ruleContext: RuleContext<MessageIds, Options>): {
        CallExpression: (node: TSESTree.CallExpression) => void;
        AssignmentExpression: (node: TSESTree.AssignmentExpression) => void;
        VariableDeclarator: (node: TSESTree.VariableDeclarator) => void;
    } {
        const services = ESLintUtils.getParserServices(ruleContext);

        const typeChecker = services.program.getTypeChecker();

        return {
            CallExpression: function(node: TSESTree.CallExpression) {
                let textArray = ancestorTextChain(node, services);
                let array = ancestorChain(node, services);
                let funcStart = getFunctionStarts2(node, services);
                if (isOfficeObject(node, typeChecker, services)) {

                    if (isOfficeFuncWriteOrRead(node, typeChecker, services) === OfficeCalls.READ) {

                        const customFunction = isCustomFunction(services, ruleContext);
    
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
                    const customFunction = isCustomFunction(services, ruleContext);

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
                    
                    const customFunction = isCustomFunction(services, ruleContext);

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

}

// const createRule = ESLintUtils.RuleCreator(
//   () => 'https://github.com/OfficeDev/Office-Addin-Scripts',
// );

// //------------------------------------------------------------------------------
// // Rule Definition
// //------------------------------------------------------------------------------

// type Options = unknown[];
// type MessageIds = 'officeReadCall';

// export = createRule<Options, MessageIds>({
//     name: 'no-office-read-calls',

//     meta: {
//         docs: {
//             description: "Prevents office api calls",
//             category: "Best Practices",
//             recommended: "warn",
//             requiresTypeChecking: true
//         },
//         type: "problem",
//         messages: {
//             officeReadCall: "No Office API read calls within Custom Functions"
//         },
//         schema: []
//     },

//     defaultOptions: [],
        
//     create(ruleContext) {
//         const services = ESLintUtils.getParserServices(ruleContext);

//         const typeChecker = services.program.getTypeChecker();

//         return {
//             CallExpression: function(node: TSESTree.CallExpression) {
//                 if (isOfficeObject(node, typeChecker, services)) {

//                     if (isOfficeFuncWriteOrRead(node, typeChecker, services) === OfficeCalls.READ) {
//                         const customFunction = isCustomFunction(services, ruleContext);
    
//                         if (customFunction) {
//                             ruleContext.report({
//                                 messageId: "officeReadCall",
//                                 loc: node.loc,
//                                 node: node
//                             });
//                         }

//                     }
//                 }
//             },

//             AssignmentExpression: function(node: TSESTree.AssignmentExpression) {
//                 if (isOfficeObject(node.right, typeChecker, services)) {
//                     const customFunction = isCustomFunction(services, ruleContext);

//                     if (customFunction) {
//                         ruleContext.report({
//                             messageId: "officeReadCall",
//                             loc: node.loc,
//                             node: node
//                         });
//                     }
//                 }
//             },

//             VariableDeclarator: function(node: TSESTree.VariableDeclarator) {
//                 if (node.init && isOfficeObject(node.init, typeChecker, services)) {
                    
//                     const customFunction = isCustomFunction(services, ruleContext);

//                     if (customFunction) {
//                         ruleContext.report({
//                             messageId: "officeReadCall",
//                             loc: node.loc,
//                             node: node
//                         });
//                     }
//                 }
//             }
//         };
//     }
// })



