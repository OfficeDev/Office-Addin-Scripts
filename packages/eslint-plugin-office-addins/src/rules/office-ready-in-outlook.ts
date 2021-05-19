import { TSESTree } from "@typescript-eslint/typescript-estree";

export = {
  name: "office-ready-in-outlook",
  meta: {
    type: <"problem" | "suggestion" | "layout">"suggestion",
    messages: {
      officeOnReadyInOutlook: "Prefer calling Office.onReady() instead of Office.initialize in Outlook applications",
    },
    docs: {
      description: "It is not a good idea to call Office.initialize in Outlook",
      category: <"Best Practices" | "Stylistic Issues" | "Variables" | "Possible Errors">"Best Practices",
      recommended: <false | "error" | "warn">false,
      url: "https://docs.microsoft.com/en-us/office/dev/add-ins/develop/initialize-add-in#initialize-with-officeonready",
    },
    schema: [],
  },
  create: function (context: any) {
    return {
      "AssignmentExpression[left.object.name='Office'][left.property.name='initialize']"(
        node: TSESTree.AssignmentExpression
      ) {
        context.report({
          node: node,
          messageId: "officeOnReadyInOutlook",
        });
      },
    };
  },
};
