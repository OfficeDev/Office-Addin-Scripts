import { TSESTree } from "@typescript-eslint/experimental-utils";
import { Reference } from "@typescript-eslint/experimental-utils/dist/ts-eslint-scope";

const getFunctions: Set<string> = new Set([
  "getAbsoluteResizedRange",
  "getActiveCell",
  "getActiveChart",
  "getActiveNotebook",
  "getActiveOutline",
  "getActivePage",
  "getActiveParagraph",
  "getActiveSection",
  "getActiveSlicer",
  "getActiveWorksheet",
  "getAncestor",
  "getAsImage",
  "getBase64Image",
  "getBase64ImageSrc",
  "getBorder",
  "getBoundingRect",
  "getById",
  "getByName",
  "getByNamespace",
  "getByTag",
  "getByTitle",
  "getByTypes",
  "getCell",
  "getCellPadding",
  "getCellProperties",
  "getColumn",
  "getColumnLabelRange",
  "getColumnProperties",
  "getColumnsAfter",
  "getColumnsBefore",
  "getCount",
  "getDataBodyRange",
  "getDataCommonPostprocess",
  "getDataHierarchy",
  "getDefault",
  "getDescendants",
  "getDimensionValues",
  "getDirectPrecedents",
  "getDocument",
  "getEntireColumn",
  "getEntireRow",
  "getExtendedRange",
  "getFilterAxisRange",
  "getFirst",
  "getFooter",
  "getHeader",
  "GetHeaderRowRange",
  "getHtml",
  "getHyperlinkRanges",
  "getImage",
  "getIntersection",
  "getInvalidCells",
  "getIsActiveCollabSession",
  "getItem",
  "getItemAt",
  "getLast",
  "getLastCell",
  "getLastColumn",
  "getLastRow",
  "getLevelParagraphs",
  "getLevelString",
  "getLocation",
  "getMergedAreas",
  "getNext",
  "getNextTextRange",
  "getOffsetRange",
  "getOffsetRangeAreas",
  "getOoxml",
  "getParagraphAfter",
  "getParagraphBefore",
  "getParagraphInfo",
  "getParentComment",
  "getPivotItems",
  "getPivotTables",
  "getPrevious",
  "getPrintArea",
  "getPrintTitleColumns",
  "getPrintTitleRows",
  "getRange",
  "getRangeAreasBySheet",
  "getRangeByIndexes",
  "getRangeEdge",
  "getRanges",
  "getReferenceId",
  "getResizedRange",
  "getRestApiId",
  "getRow",
  "getRowLabelRange",
  "getRowProperties",
  "getRowsAbove",
  "getRowsBelow",
  "getSelectedRange",
  "getSelectedRanges",
  "getSelection",
  "getSpecialCells",
  "getSpillingToRange",
  "getSpillParent",
  "GetStencilInfo",
  "getSubstring",
  "getSurroundingRegion",
  "getTables",
  "getText",
  "getTextRanges",
  "getTotalRowRange",
  "getUsedRange",
  "getUsedRangeAreas",
  "getVisibleView",
  "getWindowSize",
  "getXml",
]);

const getOrNullObjectFunctions: Set<string> = new Set([
  "getActiveChartOrNullObject",
  "getActiveNotebookOrNull",
  "getActiveOutlineOrNull",
  "getActivePageOrNull",
  "getActiveParagraphOrNull",
  "getActiveSectionOrNull",
  "getActiveSlicerOrNullObject",
  "getAncestorOrNullObject",
  "getByIdOrNullObject",
  "getCellOrNullObject",
  "getFirstOrNullObject",
  "getIntersectionOrNullObject",
  "getIntersectionOrNullObject",
  "getInvalidCellsOrNullObject",
  "getItemOrNullObject",
  "getLastOrNullObject",
  "getLocationOrNullObject",
  "getNextOrNullObject",
  "getNextTextRangeOrNullObject",
  "getParagraphAfterOrNullObject",
  "getParagraphBeforeOrNullObject",
  "getPreviousOrNullObject",
  "getPrintAreaOrNullObject",
  "getPrintTitleColumnsOrNullObject",
  "getPrintTitleRowsOrNullObject",
  "getRangeAreasOrNullObjectBySheet",
  "getRangeOrNullObject",
  "getSpecialCellsOrNullObject",
  "getSpecialCellsOrNullObject",
  "getSpillingToRangeOrNullObject",
  "getSpillParentOrNullObject",
  "getUsedRangeAreasOrNullObject",
  "getUsedRangeOrNullObject",
]);

export function isGetFunction(node: TSESTree.Node): boolean {
  return (
    node.type == TSESTree.AST_NODE_TYPES.CallExpression &&
    node.callee.type === TSESTree.AST_NODE_TYPES.MemberExpression &&
    node.callee.property.type === TSESTree.AST_NODE_TYPES.Identifier &&
    (getFunctions.has(node.callee.property.name) ||
      getOrNullObjectFunctions.has(node.callee.property.name))
  );
}

export function isGetOrNullObjectFunction(node: TSESTree.Node): boolean {
  return (
    node.type == TSESTree.AST_NODE_TYPES.CallExpression &&
    node.callee.type === TSESTree.AST_NODE_TYPES.MemberExpression &&
    node.callee.property.type === TSESTree.AST_NODE_TYPES.Identifier &&
    getOrNullObjectFunctions.has(node.callee.property.name)
  );
}

export function isLoadFunction(node: TSESTree.MemberExpression): boolean {
  return (
    node.property.type === TSESTree.AST_NODE_TYPES.Identifier &&
    node.property.name === "load"
  );
}

export function isContextSyncIdentifier(node: TSESTree.Identifier): boolean {
  return (
    node.name === "context" &&
    node.parent?.type === TSESTree.AST_NODE_TYPES.MemberExpression &&
    node.parent?.parent?.type === TSESTree.AST_NODE_TYPES.CallExpression &&
    node.parent?.property.type === TSESTree.AST_NODE_TYPES.Identifier &&
    node.parent?.property.name === "sync"
  );
}

export type OfficeApiReference = {
  operation: "Read" | "Load" | "Write" | "Sync";
  reference: Reference;
};

export function isLoadReference(node: TSESTree.Identifier) {
  return (
    node.parent &&
    node.parent.type === TSESTree.AST_NODE_TYPES.MemberExpression &&
    isLoadFunction(node.parent)
  );
}
