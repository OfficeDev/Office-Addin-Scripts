import { TSESTree, ESLintUtils } from "@typescript-eslint/experimental-utils";
import { REPO_URL, isCustomFunction, isOfficeObject, isOfficeFuncWriteOrRead, OfficeCalls, isHelperFunc, getStartOfFunction, getFunctionDeclarations, superNodeMe, reportIfCalledFromCustomFunction} from './utils'
import { RuleContext, RuleMetaDataDocs, RuleMetaData  } from '@typescript-eslint/experimental-utils/dist/ts-eslint';
import ts from 'typescript';

/**
 * @fileoverview Prevents office api calls
 * @author Artur Tarasenko (artarase)
 */
"use strict";

//------------------------------------------------------------------------------
// Rule Definition
//------------------------------------------------------------------------------

type Options = unknown[];
type MessageIds = 'officeWriteCall';

export = {
    name: 'no-office-write-calls',

    meta: {
        docs: {
            description: "Prevents office write api calls",
            category: <RuleMetaDataDocs["category"]>"Best Practices",
            recommended: <RuleMetaDataDocs["recommended"]>"error",
            requiresTypeChecking: true,
            url: REPO_URL
        },
        type: <RuleMetaData<MessageIds>["type"]> "problem",
        messages: <Record<MessageIds, string>> {
            officeWriteCall: "No Office API write calls within Custom Functions"
        },
        schema: []
    },

    create: function(ruleContext: RuleContext<MessageIds, Options>): {
        CallExpression: (node: TSESTree.CallExpression) => void;
        AssignmentExpression: (node: TSESTree.AssignmentExpression) => void;
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
                    if (isOfficeFuncWriteOrRead(node, typeChecker, services) === OfficeCalls.WRITE) {
                        if (isCustomFunction(node, services)) {
                            ruleContext.report({
                                messageId: "officeWriteCall",
                                loc: node.loc,
                                node: node
                            });
                        }

                        const functionStart = getStartOfFunction(node, services);
                        if (functionStart) {
                            reportIfCalledFromCustomFunction(functionStart, 
                                ruleContext, 
                                helperFuncToHelperFuncMap, 
                                helperFuncToMentionsMap, 
                                officeCallingFuncs
                            );
                        }
                    }
                } else if (isHelperFunc(node, typeChecker, services)) {
                    const functionDeclarations = getFunctionDeclarations(node, typeChecker, services);

                    if (functionDeclarations && functionDeclarations.length > 0) {
                        superNodeMe(functionDeclarations, helperFuncToHelperFuncMap);

                        if(isCustomFunction(node, services)) {
                            if (functionDeclarations.some((declaration) => {
                                return officeCallingFuncs.has(declaration);
                            })) {
                                ruleContext.report({
                                    messageId: "officeWriteCall",
                                    loc: node.loc,
                                    node: node
                                });
                            } else {
                                helperFuncToMentionsMap.set(functionDeclarations[0], 
                                    (helperFuncToMentionsMap.get(functionDeclarations[0]) || []).concat({
                                        messageId: "officeWriteCall",
                                        loc: node.loc,
                                        node: node
                                    })
                                );
                            }
                        }

                        const functionStart = getStartOfFunction(node, services);
                        
                        if (functionStart) {
                            helperFuncToHelperFuncMap.set(
                                functionDeclarations[0], 
                                (helperFuncToHelperFuncMap.get(functionDeclarations[0]) || new Set<ts.Node>([])).add(functionStart)
                            );
                            reportIfCalledFromCustomFunction(functionStart, 
                                ruleContext, 
                                helperFuncToHelperFuncMap, 
                                helperFuncToMentionsMap
                            );
                        }
                    }
                }
            },

            AssignmentExpression: function(node: TSESTree.AssignmentExpression) {
                if (isOfficeObject(node.left, typeChecker, services)
                && isCustomFunction(node, services)) {
                    ruleContext.report({
                        messageId: "officeWriteCall",
                        loc: node.loc,
                        node: node
                    });
                }
            }
        };
    }
}