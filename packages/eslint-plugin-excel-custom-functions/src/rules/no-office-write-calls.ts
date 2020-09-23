import { TSESTree, ESLintUtils, ParserServices } from "@typescript-eslint/experimental-utils";
import { REPO_URL, callExpressionAnalysis, assignmentExpressionAnalysis} from './utils'
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
        const services: ParserServices = ESLintUtils.getParserServices(ruleContext);
        const typeChecker: ts.TypeChecker = services.program.getTypeChecker();

        //Registry of all functions that use Office API calls (regardless if CF or not)
        let officeCallingFuncs = new Set<ts.Node>(); 

        //Registry of all times user-created helper functions are used in CF (regardless if they call Office API calls or not)
        let helperFuncToMentionsMap = new Map<ts.Node, Array<{messageId: MessageIds, loc: TSESTree.SourceLocation, node: TSESTree.Node}>>(); 

        //Mapping of all helper funcs to the functions they get called within
        let helperFuncToHelperFuncMap = new Map<ts.Node, Set<ts.Node>>();

        return {
            CallExpression: function(node: TSESTree.CallExpression) {
                callExpressionAnalysis(node, 
                    services, 
                    typeChecker, 
                    ruleContext, 
                    officeCallingFuncs, 
                    helperFuncToMentionsMap, 
                    helperFuncToHelperFuncMap, 
                    true
                );
            },

            AssignmentExpression: function(node: TSESTree.AssignmentExpression) {
                assignmentExpressionAnalysis(node, ruleContext, services, typeChecker, true);
            },
        };
    }
}