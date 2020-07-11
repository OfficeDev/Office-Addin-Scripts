"use strict";
exports.__esModule = true;
var experimental_utils_1 = require("@typescript-eslint/experimental-utils");
var utils_1 = require("./utils");
/**
 * @fileoverview Prevents office api calls
 * @author Artur Tarasenko (artarase)
 */
"use strict";
//------------------------------------------------------------------------------
// Rule Definition
//------------------------------------------------------------------------------
// let excelRunToContextMap: Map<TSESTree.Node, TSESTree.Identifier> = new Map<TSESTree.Node, TSESTree.Identifier>();
// let contextToExcelRunMap: Map<TSESTree.Node, TSESTree.Node> = new Map<TSESTree.Node, TSESTree.Node>();
var excelRunToContextMap = new Map();
var contextToExcelRunMap = new Map();
function isInExcelRun(node) {
    if (excelRunToContextMap.has(node)) {
        return node;
    }
    else {
        return node.parent ? isInExcelRun(node.parent) : undefined;
    }
}
exports["default"] = utils_1.createRule({
    name: __filename,
    meta: {
        docs: {
            description: "Prevents office api calls",
            category: "Possible Errors",
            recommended: "error"
        },
        fixable: undefined,
        schema: [
        // fill in your schema
        ],
        type: "problem",
        messages: {
            contextSync: "No context.sync() calls within Custom Functions"
        }
    },
    defaultOptions: [],
    create: function (ruleContext) {
        return {
            CallExpression: function (node) {
                if (utils_1.isOfficeBoilerplate(node)) {
                    if (node.arguments[0].type == experimental_utils_1.AST_NODE_TYPES.FunctionExpression
                        && node.arguments[0].params.length > 0
                        && node.arguments[0].params[0].type == experimental_utils_1.AST_NODE_TYPES.Identifier) {
                        excelRunToContextMap.set(node, node.arguments[0].params[0]);
                        contextToExcelRunMap.set(node.arguments[0].params[0], node);
                    }
                }
            },
            Identifier: function (node) {
                var excelRunNode = isInExcelRun(node);
                var originalContext;
                if (!!excelRunNode && excelRunToContextMap.has(excelRunNode)) {
                    originalContext = excelRunToContextMap.get(excelRunNode);
                    if ((originalContext === null || originalContext === void 0 ? void 0 : originalContext.name) == node.name) {
                        contextToExcelRunMap.set(node, excelRunNode);
                    }
                }
            },
            "MemberExpression > Identifier:exit": function (node) {
                if (contextToExcelRunMap.has(node)
                    && node.parent && node.parent.type == experimental_utils_1.AST_NODE_TYPES.MemberExpression
                    && (node.parent).property.type == experimental_utils_1.AST_NODE_TYPES.Identifier
                    && (node.parent).property.name == "sync") {
                    ruleContext.report({
                        messageId: "contextSync",
                        loc: node.parent.loc,
                        node: node.parent
                    });
                }
            }
        };
    }
});
