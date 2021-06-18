import { TSESTree } from "@typescript-eslint/experimental-utils";
import {
  Reference,
  Scope,
  Variable,
} from "@typescript-eslint/experimental-utils/dist/ts-eslint-scope";

export = {
  name: "load-object-before-read",
  meta: {
    type: <"problem" | "suggestion" | "layout">"problem",
    messages: {
      loadBeforeRead:
        "An explicit load call on '{{name}}' for '{{loadValue}}' needs to be made before reading a proxy object",
    },
    docs: {
      description:
        "Before you can read the properties of a proxy object, you must explicitly load the properties",
      category: <
        "Best Practices" | "Stylistic Issues" | "Variables" | "Possible Errors"
      >"Possible Errors",
      recommended: <false | "error" | "warn">false,
      url: "https://docs.microsoft.com/en-us/office/dev/add-ins/develop/application-specific-api-model?view=powerpoint-js-1.1#load",
    },
    schema: [],
  },
  create: function (context: any) {
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

    function findPropertiesRead(node: TSESTree.Node | undefined): string {
      let propertyName = ""; // Will be a string combined with '/' for the case of navigation properties
      while (node) {
        if (
          node.type === TSESTree.AST_NODE_TYPES.MemberExpression &&
          node.property.type === TSESTree.AST_NODE_TYPES.Identifier
        ) {
          propertyName += node.property.name + "/";
        }
        node = node.parent;
      }
      return propertyName.slice(0, -1);
    }

    function isLoadFunction(node: TSESTree.MemberExpression): boolean {
      return (
        node.property.type === TSESTree.AST_NODE_TYPES.Identifier &&
        node.property.name === "load"
      );
    }

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

    function getPropertyName(node: TSESTree.MemberExpression): string {
      if (
        node.parent?.type === TSESTree.AST_NODE_TYPES.CallExpression &&
        node.parent.arguments[0].type === TSESTree.AST_NODE_TYPES.Literal
      ) {
        return node.parent.arguments[0].value as string;
      }
      return "error in getPropertyName";
    }

    function isInsideWriteStatement(node: TSESTree.Node): boolean {
      while (node.parent) {
        node = node.parent;
        if (node.type === TSESTree.AST_NODE_TYPES.AssignmentExpression)
          return true;
      }
      return false;
    }

    function findLoadBeforeRead(scope: Scope) {
      scope.variables.forEach((variable: Variable) => {
        let loadLocation: Map<string, number> = new Map<string, number>();
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

          if (node.parent?.type === TSESTree.AST_NODE_TYPES.MemberExpression) {
            if (isLoadFunction(node.parent)) {
              // In case it is a load function
              loadLocation.set(getPropertyName(node.parent), node.range[1]);
              return;
            }
          }

          const propertyName: string | undefined = findPropertiesRead(
            node.parent
          );
          if (!propertyName) {
            // There is no property
            return;
          }

          if (
            loadLocation.has(propertyName) && // If reference came after load, return
            node.range[1] > (loadLocation.get(propertyName) ?? 0)
          ) {
            return;
          }

          if (isInsideWriteStatement(node)) {
            // Return in case it a write, ie, not read statment
            return;
          }

          context.report({
            node: node,
            messageId: "loadBeforeRead",
            data: { name: node.name, loadValue: propertyName },
          });
        });
      });
      scope.childScopes.forEach(findLoadBeforeRead);
    }

    return {
      Program() {
        findLoadBeforeRead(context.getScope());
      },
    };
  },
};
