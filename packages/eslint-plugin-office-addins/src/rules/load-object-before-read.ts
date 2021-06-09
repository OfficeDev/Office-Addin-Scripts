import { TSESTree } from "@typescript-eslint/typescript-estree";
import {
  Reference,
  Scope,
  Variable,
} from "@typescript-eslint/experimental-utils/dist/ts-eslint-scope";
import { isConstructorDeclaration } from "typescript";

export = {
  name: "load-object-before-read",
  meta: {
    type: <"problem" | "suggestion" | "layout">"problem",
    messages: {
      loadBeforeRead: "An explicit load call on '{{name}}' for '{{loadValue}}' needs to be made before reading a proxy object",
    },
    docs: {
      description: 
        "Before you can read the properties of a proxy object, you must explicitly load the properties",
      category: <
        "Best Practices" | "Stylistic Issues" | "Variables" | "Possible Errors"
      >"Possible Errors",
      recommended: <false | "error" | "warn">false,
      url: 
        "https://docs.microsoft.com/en-us/office/dev/add-ins/develop/application-specific-api-model?view=powerpoint-js-1.1#load",
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

    function getPropertyThatHadToBeLoaded(node: TSESTree.Node): string | undefined {
      if(node.parent?.type === TSESTree.AST_NODE_TYPES.MemberExpression
        && node.parent?.property.type === TSESTree.AST_NODE_TYPES.Identifier) {
        return node.parent.property.name;
      }
      return undefined;
    }

    function isLoadFunction(node: TSESTree.Node): boolean {
      return (node.parent?.type === TSESTree.AST_NODE_TYPES.MemberExpression
        && node.parent.property.type === TSESTree.AST_NODE_TYPES.Identifier
        && node.parent.property.name === "load");
    }

    function callsGetAPIFunction(node: TSESTree.Identifier): boolean {
      return getFunctions.has(node.name);
    }

    function isVariableDeclaration(node: TSESTree.Node): boolean {
      if(node.parent?.type === TSESTree.AST_NODE_TYPES.VariableDeclarator) {
          return true;
      }
      return false;
    }

    function isGetVariableDeclaration(node: TSESTree.Node): boolean {
      if(node.parent?.type === TSESTree.AST_NODE_TYPES.VariableDeclarator
        && node.parent.init?.type === TSESTree.AST_NODE_TYPES.CallExpression
        && node.parent.init.callee.type === TSESTree.AST_NODE_TYPES.MemberExpression
        && node.parent.init.callee.property.type === TSESTree.AST_NODE_TYPES.Identifier) {
          if(callsGetAPIFunction(node.parent.init.callee.property)) {
            return true;
          }
      }
      return false;
    }

    function isAssignmentExpression(node: TSESTree.Node): boolean {
      if(node.parent?.type === TSESTree.AST_NODE_TYPES.AssignmentExpression) {
          return true;
      }
      return false;
    }

    function isGetAssignmentExpression(node: TSESTree.Node): boolean {
      if(node.parent?.type === TSESTree.AST_NODE_TYPES.AssignmentExpression
        && node.parent.right.type === TSESTree.AST_NODE_TYPES.CallExpression
        && node.parent.right.callee.type === TSESTree.AST_NODE_TYPES.MemberExpression
        && node.parent.right.callee.property.type === TSESTree.AST_NODE_TYPES.Identifier) {
          if(callsGetAPIFunction(node.parent.right.callee.property)) {
            return true;
          }
      }
      return false;
    }

    function getLoadedPropertyName(node: TSESTree.Node): string {
      if(node.parent?.parent?.type === TSESTree.AST_NODE_TYPES.CallExpression
        && node.parent.parent.arguments[0].type === TSESTree.AST_NODE_TYPES.Literal) {
          return node.parent.parent.arguments[0].value as string;
        }
      return "error in getLoadedPropertyName";
    }

    function findLoadBeforeRead(scope: Scope) {
      scope.variables.forEach((variable: Variable) => {
        let loadLocation: Map <string, number> = new Map<string, number>();
        let getFound: boolean = false;
        variable.references.forEach((reference: Reference) => {
          const node: TSESTree.Node = reference.identifier;
            if (isVariableDeclaration(node)
              || isAssignmentExpression(node)) {
              if (isGetVariableDeclaration(node)
                || isGetAssignmentExpression(node)) {
                  getFound = true;
                  return;
              } else {
                getFound = false;
              }
            }
          
          if(!getFound) {
            return;
          }

          if (isLoadFunction(node)) {
            loadLocation.set(getLoadedPropertyName(node), node.range[1]);
            return;
          }

          // If reference came after load 
          const propertyName: string | undefined = getPropertyThatHadToBeLoaded(node);
          if (!propertyName) {
            return;
          }

          if (loadLocation.has(propertyName)
            && (node.range[1] > (loadLocation.get(propertyName) ?? 0))) {
              return;
          }

          context.report({
            node: node,
            messageId: "loadBeforeRead",
            data: {name: node.name, loadValue: propertyName}
          });
        });
      });
      scope.childScopes.forEach(findLoadBeforeRead);
    }

    return {
      Program() {
        findLoadBeforeRead(context.getScope());
      }
    }
  },
};
