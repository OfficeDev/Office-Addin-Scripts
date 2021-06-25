import { TSESTree } from "@typescript-eslint/experimental-utils";
import {
  Reference,
  Scope,
  Variable,
} from "@typescript-eslint/experimental-utils/dist/ts-eslint-scope";
//import { isGetFunction, isLoadFunction } from "../utils";

const getFunctions: Set<string> = new Set([
  "getDataCommonPostprocess",
  "getRange",
  "getRangeOrNullObject",
  "getActiveCell",
  "getActiveChart",
  "getActiveChartOrNullObject",
  "getActiveSlicer",
  "getActiveSlicerOrNullObject",
  "getIsActiveCollabSession",
  "getSelectedRange",
  "getSelectedRanges",
  "getCell",
  "getNext",
  "getNextOrNullObject",
  "getPrevious",
  "getPreviousOrNullObject",
  "getRangeByIndexes",
  "getRanges",
  "getUsedRange",
  "getUsedRangeOrNullObject",
  "getActiveWorksheet",
  "getCount",
  "getFirst",
  "getFirstOrNullObject",
  "getItem",
  "getItemOrNullObject",
  "getLast",
  "getLocation",
  "getLocationOrNullObject",
  "getAbsoluteResizedRange",
  "getBoundingRect",
  "getCellProperties",
  "getColumn",
  "getColumnProperties",
  "getColumnsAfter",
  "getColumnsBefore",
  "getDirectPrecedents",
  "getEntireColumn",
  "getEntireRow",
  "getExtendedRange",
  "getImage",
  "getIntersection",
  "getIntersectionOrNullObject",
  "getLastCell",
  "getLastColumn",
  "getLastRow",
  "getMergedAreas",
  "getOffsetRange",
  "getPivotTables",
  "getRangeEdge",
  "getResizedRange",
  "getRow",
  "getRowProperties",
  "getRowsAbove",
  "getRowsBelow",
  "getSpecialCells",
  "getSpecialCellsOrNullObject",
  "getSpillParent",
  "getSpillParentOrNullObject",
  "getSpillingToRange",
  "getSpillingToRangeOrNullObject",
  "getSurroundingRegion",
  "getTables",
  "getVisibleView",
  "getEntireColumn",
  "getEntireRow",
  "getIntersection",
  "getIntersectionOrNullObject",
  "getOffsetRangeAreas",
  "getSpecialCells",
  "getSpecialCellsOrNullObject",
  "getUsedRangeAreas",
  "getUsedRangeAreasOrNullObject",
  "getRangeAreasBySheet",
  "getRangeAreasOrNullObjectBySheet",
  "getItemAt",
  "getText",
  "getDataBodyRange",
  "GetHeaderRowRange",
  "getSubstring",
  "getTotalRowRange",
  "getInvalidCells",
  "getInvalidCellsOrNullObject",
  "getDimensionValues",
  "getByNamespace",
  "getXml",
  "getColumnLabelRange",
  "getDataHierarchy",
  "getFilterAxisRange",
  "getPivotItems",
  "getRowLabelRange",
  "getDefault",
  "getPrintArea",
  "getPrintAreaOrNullObject",
  "getPrintTitleColumns",
  "getPrintTitleColumnsOrNullObject",
  "getPrintTitleRows",
  "getPrintTitleRowsOrNullObject",
  "getParentComment",
  "getAsImage",
  "getActivePage",
  "GetStencilInfo",
  "getBase64Image",
  "getHtml",
  "getParagraphInfo",
  "getByTitle",
  "getRestApiId",
  "getByName",
  "getWindowSize",
  "getActiveSectionOrNull",
  "getActiveSection",
  "getActiveParagraphOrNull",
  "getActiveParagraph",
  "getActivePageOrNull",
  "getActivePage",
  "getActiveOutlineOrNull",
  "getActiveOutline",
  "getActiveNotebookOrNull",
  "getActiveNotebook",
  "getCellPadding",
  "getBorder",
  "getParagraphBefore",
  "getParagraphBeforeOrNullObject",
  "getParagraphAfterOrNullObject",
  "getParagraphAfter",
  "getCellOrNullObject",
  "getHeader",
  "getFooter",
  "getTextRanges",
  "getOoxml",
  "getNextTextRangeOrNullObject",
  "getNextTextRange",
  "getHyperlinkRanges",
  "getLastOrNullObject",
  "getDescendants",
  "getAncestorOrNullObject",
  "getAncestor",
  "getByIdOrNullObject",
  "getById",
  "getLevelString",
  "getLevelParagraphs",
  "getBase64ImageSrc",
  "getSelection",
  "getByTypes",
  "getByTag",
  "getReferenceId",
  "getDocument",
]);
function callsGetAPIFunction(node: TSESTree.Identifier): boolean {
  return getFunctions.has(node.name);
}

function isGetFunction(node: TSESTree.Expression): boolean {
  if (
    node.type == TSESTree.AST_NODE_TYPES.CallExpression &&
    node.callee.type === TSESTree.AST_NODE_TYPES.MemberExpression &&
    node.callee.property.type === TSESTree.AST_NODE_TYPES.Identifier
  ) {
    if (callsGetAPIFunction(node.callee.property)) {
      return true;
    }
  }
  return false;
}
function isLoadFunction(node: TSESTree.MemberExpression): boolean {
  return (
    node.property.type === TSESTree.AST_NODE_TYPES.Identifier &&
    node.property.name === "load"
  );
}


export = {
  name: "no-empty-load",
  meta: {
    type: <"problem" | "suggestion" | "layout">"problem",
    messages: {
      emptyLoad:
        "Calling load without any argument can slow down your add-in",
    },
    docs: {
      description:
        "Calling load without any argument can cause load unneeded data and slow down your add-in",
      category: <
        "Best Practices" | "Stylistic Issues" | "Variables" | "Possible Errors"
      >"Best Practices",
      recommended: <false | "error" | "warn">false,
      url: "https://docs.microsoft.com/en-us/office/dev/add-ins/develop/application-specific-api-model?view=powerpoint-js-1.1#calling-load-without-parameters-not-recommended",
    },
    schema: [],
  },
  create: function (context: any) {
    function isEmptyLoad(node: TSESTree.MemberExpression): boolean {
      if(isLoadFunction(node)) {
        //console.log("Checking for empty load");
        //console.log(node.parent);
        if(node.parent?.type == TSESTree.AST_NODE_TYPES.CallExpression) {
          //console.log("Inside first if");
          //console.log(node.parent.arguments);
          //console.log(node.parent.arguments === []);
          if (node.parent.arguments.length === 0) {
            //console.log("Retuned true");
            return true;
          }
        }
      }
      return false;
    }

    function findEmptyLoad(scope: Scope) {
      scope.variables.forEach((variable: Variable) => {
        let getFound: boolean = false;
        variable.references.forEach((reference: Reference) => {
          const node: TSESTree.Node = reference.identifier;

          if (
            node.parent?.type === TSESTree.AST_NODE_TYPES.VariableDeclarator
          ) {
            getFound = false; // In case of reassignment
            if (node.parent.init && isGetFunction(node.parent.init)) {
              getFound = true;
              return;
            }
          }

          if (
            node.parent?.type === TSESTree.AST_NODE_TYPES.AssignmentExpression
          ) {
            getFound = false; // In case of reassignment
            if (isGetFunction(node.parent.right)) {
              getFound = true;
              return;
            }
          }

          if (!getFound) {
            // If reference was not related to a previous get
            return;
          }

          if (node.parent?.type === TSESTree.AST_NODE_TYPES.MemberExpression
            && isEmptyLoad(node.parent)) {
              context.report({
                node: node,
                messageId: "emptyLoad",
              });
          }
        });
      });
      scope.childScopes.forEach(findEmptyLoad);
    }

    return {
      Program() {
        findEmptyLoad(context.getScope());
      },
    };
  },
};
