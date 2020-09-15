// import { TSESTree, ESLintUtils } from "@typescript-eslint/experimental-utils";
// //import * as sm from "@typescript-eslint/scope-manager";
// import { getCustomFunction, isOfficeObject, isOfficeFuncWriteOrRead, OfficeCalls } from './utils'
// import { RuleContext, RuleMetaDataDocs, RuleMetaData } from '@typescript-eslint/experimental-utils/dist/ts-eslint';
// import ts from 'typescript';

// /**
//  * @fileoverview Prevents office api calls
//  * @author Artur Tarasenko (artarase)
//  */
// "use strict";

// const createRule = ESLintUtils.RuleCreator(
//   () => 'https://github.com/OfficeDev/Office-Addin-Scripts',
// );

// //------------------------------------------------------------------------------
// // Rule Definition
// //------------------------------------------------------------------------------

// type Options = unknown[];
// type MessageIds = 'officeWriteCall';

// export = {
//     name: 'no-office-write-calls',

//     meta: {
//         docs: {
//             description: "Prevents office api calls",
//             category: <RuleMetaDataDocs["category"]>"Best Practices",
//             recommended: <RuleMetaDataDocs["recommended"]>"error",
//             requiresTypeChecking: true,
//             url: 'https://github.com/OfficeDev/Office-Addin-Scripts'
//         },
//         type: <RuleMetaData<MessageIds>["type"]> "problem",
//         messages: <Record<MessageIds, string>> {
//             officeWriteCall: "No Office API write calls within Custom Functions"
//         },
//         schema: []
//     },

//     create: function(ruleContext: RuleContext<MessageIds, Options>): {
//         CallExpression: (node: TSESTree.CallExpression) => void;
//         AssignmentExpression: (node: TSESTree.AssignmentExpression) => void;
//     } {

//         const services = ESLintUtils.getParserServices(ruleContext);

//         const typeChecker = services.program.getTypeChecker();

//         const sourceFiles = services.program.getSourceFiles();

//         let betterSourceFiles = [];

//         for (let i = 0; i < sourceFiles.length; i++) {
//             if(!sourceFiles[i].isDeclarationFile) {
//                 betterSourceFiles.push(sourceFiles[i]);
//             }
//         }

//         let functionMap = new Map<ts.Node, OfficeCalls>();

//         return {
//             CallExpression: function(node: TSESTree.CallExpression) {
//                 const customFunction = getCustomFunction(services, ruleContext);

//                 if (customFunction) {
//                     if(isOfficeObject(node, typeChecker, services)) {
//                         if (isOfficeFuncWriteOrRead(node, typeChecker, services) === OfficeCalls.WRITE) {
//                             ruleContext.report({
//                                 messageId: "officeWriteCall",
//                                 loc: node.loc,
//                                 node: node
//                             });
//                         }
//                     }
//                 } else 
//             },

//             AssignmentExpression: function(node: TSESTree.AssignmentExpression) {
//                 if (isOfficeObject(node.left, typeChecker, services)) {
                    
//                     const customFunction = getCustomFunction(services, ruleContext);

//                     if (customFunction) {
//                         ruleContext.report({
//                             messageId: "officeWriteCall",
//                             loc: node.loc,
//                             node: node
//                         });
//                     }
//                 }
//             }
//         };
//     }

// }
// // const createRule = ESLintUtils.RuleCreator(
// //   () => 'https://github.com/OfficeDev/Office-Addin-Scripts',
// // );

// // //------------------------------------------------------------------------------
// // // Rule Definition
// // //------------------------------------------------------------------------------

// // type Options = unknown[];
// // type MessageIds = 'officeWriteCall';

// // export = createRule<Options, MessageIds>({
// //     name: 'no-office-write-calls',

// //     meta: {
// //         docs: {
// //             description: "Prevents office api calls",
// //             category: "Best Practices",
// //             recommended: "error",
// //             requiresTypeChecking: true
// //         },
// //         type: "problem",
// //         messages: {
// //             officeWriteCall: "No Office API write calls within Custom Functions"
// //         },
// //         schema: []
// //     },

// //     defaultOptions: [],
        
// //     create(ruleContext) {

// //         const services = ESLintUtils.getParserServices(ruleContext);

// //         const typeChecker = services.program.getTypeChecker();

// //         return {
// //             CallExpression: function(node: TSESTree.CallExpression) {

// //                 if(isOfficeObject(node, typeChecker, services)) {

// //                     if (isOfficeFuncWriteOrRead(node, typeChecker, services) === OfficeCalls.WRITE) {
// //                         const customFunction = getCustomFunction(services, ruleContext);

// //                         if (customFunction) {
// //                             ruleContext.report({
// //                                 messageId: "officeWriteCall",
// //                                 loc: node.loc,
// //                                 node: node
// //                             });
// //                         }

// //                     }
// //                 }
// //             },

// //             AssignmentExpression: function(node: TSESTree.AssignmentExpression) {
// //                 if (isOfficeObject(node.left, typeChecker, services)) {
                    
// //                     const customFunction = getCustomFunction(services, ruleContext);

// //                     if (customFunction) {
// //                         ruleContext.report({
// //                             messageId: "officeWriteCall",
// //                             loc: node.loc,
// //                             node: node
// //                         });
// //                     }
// //                 }
// //             }
// //         };
// //     }
// // })



