import { TSESTree, ESLintUtils } from "@typescript-eslint/experimental-utils";
import { getCustomFunction, isOfficeObject, isOfficeFuncWriteOrRead, OfficeCalls, getFunctionStarts2, isHelperFunc, bubbleUpNewCallingFuncs, getFunctionDeclarations, superNodeMe } from './utils'
import { RuleContext, RuleMetaDataDocs, RuleMetaData  } from '@typescript-eslint/experimental-utils/dist/ts-eslint';
import ts from 'typescript';

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

        //Registry of all functions that use Office API calls (regardless if CF or not)
        let officeCallingFuncs = new Set<ts.Node>(); 

        //Registry of all times user-created helper functions are used in CF (regardless if they call Office API calls or not)
        let helperFuncToMentionsMap = new Map<ts.Node, Array<{messageId: MessageIds, loc: TSESTree.SourceLocation, node: TSESTree.Node}>>(); 

        //Mapping of all helper funcs to the functions they get called within
        let helperFuncToHelperFuncMap = new Map<ts.Node, Set<ts.Node>>();


        return {
            CallExpression: function(node: TSESTree.CallExpression) {
                if (isOfficeObject(node, typeChecker, services)) {
                    if (isOfficeFuncWriteOrRead(node, typeChecker, services) === OfficeCalls.READ) {

                        if (getCustomFunction(services, ruleContext)) {
                            ruleContext.report({
                                messageId: "officeReadCall",
                                loc: node.loc,
                                node: node
                            });
                        }

                        const functionStarts = getFunctionStarts2(node, services);
                        functionStarts.forEach((functionStart) => {
                            const bubbledUp = bubbleUpNewCallingFuncs(functionStart, helperFuncToHelperFuncMap);
                            bubbledUp.forEach((newEntry) => {
                                officeCallingFuncs.add(newEntry);
                                helperFuncToMentionsMap.get(newEntry)?.forEach((mention) => {
                                    ruleContext.report(mention);
                                });
                                helperFuncToMentionsMap.delete(newEntry);
                            });
                        });
                    }
                } else if (isHelperFunc(node, typeChecker, services)) {

                    const customFunction = getCustomFunction(services, ruleContext);
                    const functionStarts = getFunctionStarts2(node, services);
                    const functionDeclarations = getFunctionDeclarations(node, typeChecker, services);
                    if (functionDeclarations) {
                        superNodeMe(functionDeclarations, helperFuncToHelperFuncMap);
                    }

                    if (functionDeclarations && functionDeclarations.length > 0) {
                        if(customFunction) {
                            let newMentionsArray = helperFuncToMentionsMap.get(functionDeclarations[0]);
                            helperFuncToMentionsMap.set(functionDeclarations[0], 
                                newMentionsArray ? 
                                newMentionsArray.concat({
                                    messageId: "officeReadCall",
                                    loc: node.loc,
                                    node: node
                                }) :
                                [{
                                    messageId: "officeReadCall",
                                    loc: node.loc,
                                    node: node
                                }]
                            );
                        }
                        
                        functionStarts.forEach((functionStart) => {
                            let newHelperFuncSet = helperFuncToHelperFuncMap.get(functionStart);
                            if (!newHelperFuncSet) {
                                newHelperFuncSet = new Set<ts.Node>([]);
                            }
                            helperFuncToHelperFuncMap.set(functionStart, 
                                newHelperFuncSet.add(functionDeclarations[0])
                            );
                        });
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
//                         const customFunction = getCustomFunction(services, ruleContext);
    
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
//                     const customFunction = getCustomFunction(services, ruleContext);

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
                    
//                     const customFunction = getCustomFunction(services, ruleContext);

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



