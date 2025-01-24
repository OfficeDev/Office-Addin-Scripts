import { ESLintUtils, TSESTree } from "@typescript-eslint/utils";

export default ESLintUtils.RuleCreator(
  () =>
    "https://docs.microsoft.com/office/dev/add-ins/develop/initialize-add-in#initialize-with-officeonready",
)({
  name: "no-office-initialize",
  meta: {
    type: "suggestion",
    messages: {
      noOfficeInitialize:
        "Office.onReady() is preferred over Office.initialize.",
    },
    docs: {
      description: "Office.onReady() is more flexible than Office.initialize.",
    },
    schema: [],
  },
  create: function (context) {
    return {
      "AssignmentExpression[left.object.name='Office'][left.property.name='initialize']"(
        node: TSESTree.AssignmentExpression,
      ) {
        context.report({
          node: node,
          messageId: "noOfficeInitialize",
        });
      },
    };
  },
  defaultOptions: [],
});
