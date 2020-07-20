import { TSESTree, ESLintUtils, TSESLint, AST_NODE_TYPES } from "@typescript-eslint/experimental-utils";
import { isCallSignatureDeclaration, isIdentifier } from "typescript";

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

// let excelRunToContextMap: Map<TSESTree.Node, TSESTree.Identifier> = new Map<TSESTree.Node, TSESTree.Identifier>();
// let contextToExcelRunMap: Map<TSESTree.Node, TSESTree.Node> = new Map<TSESTree.Node, TSESTree.Node>();
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
        return {
            CallExpression: function(node: TSESTree.CallExpression) {
                ruleContext.report({
                    messageId: "contextSync",
                    loc: node.loc,
                    node: node
                });
                // if(isOfficeBoilerplate(node)) {
                //     // if(node.arguments[0].type == AST_NODE_TYPES.FunctionExpression
                //     //     && node.arguments[0].params.length > 0
                //     //     && node.arguments[0].params[0].type == AST_NODE_TYPES.Identifier) {
                //     //     excelRunToContextMap.set(node, node.arguments[0].params[0]);
                //     //     contextToExcelRunMap.set(node.arguments[0].params[0], node);
                //     // }
                // }
            },

            Identifier: function(node: TSESTree.Identifier) {
                // let excelRunNode = isInExcelRun(node);
                // let originalContext: TSESTree.Identifier | undefined;

                // if (!!excelRunNode && excelRunToContextMap.has(excelRunNode)) {
                //     originalContext = excelRunToContextMap.get(excelRunNode);
                //     if(originalContext?.name == node.name) {
                //         contextToExcelRunMap.set(node, excelRunNode);
                //         console.log(excelRunToContextMap);
                //         console.log(contextToExcelRunMap);
                //     }
                // }
                ruleContext.report({
                    messageId: "contextSync",
                    loc: node.loc,
                    node: node
                });
            },

            // "Identifier:exit": function(node: TSESTree.Identifier) {
            //     if (contextToExcelRunMap.has(node)
            //         && node.parent && node.parent.type == AST_NODE_TYPES.MemberExpression 
            //         && (<TSESTree.MemberExpression>(node.parent)).property.type == AST_NODE_TYPES.Identifier
            //         && (<TSESTree.Identifier>(<TSESTree.MemberExpression>(node.parent)).property).name == "sync") {
            //             ruleContext.report({
            //                 messageId: "contextSync",
            //                 loc: node.parent.loc,
            //                 node: node.parent
            //             });
            //     }
            // }
        };
    }
})



