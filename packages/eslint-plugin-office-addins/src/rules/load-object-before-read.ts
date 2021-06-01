import { TSESTree } from "@typescript-eslint/typescript-estree";

export = {
  name: "load-object-before-read",
  meta: {
    type: <"problem" | "suggestion" | "layout">"problem",
    messages: {
      loadBeforeRead: "An explicit load call needs to be made before reading a proxu object",
    },
    docs: {
      description: "Before you can read the properties of a proxy object, you must explicitly load the properties",
      category: <"Best Practices" | "Stylistic Issues" | "Variables" | "Possible Errors">"Possible Errors",
      recommended: <false | "error" | "warn">false,
      url: "https://docs.microsoft.com/en-us/office/dev/add-ins/develop/application-specific-api-model?view=powerpoint-js-1.1#load",
    },
    schema: [],
  },
  create: function (context: any) {
    return {
      "MemberExpression[left.object.name='Office'][left.property.name='initialize']"(
        node: TSESTree.MemberExpression
      ) {
        context.report({
          node: node,
          messageId: "loadBeforeRead",
        });
      },
    };
  },
};

/*
Locating Variable:
VariableDeclarator :matches Identifier.name (this is the variable name) then check for get call expression

Locating the get functions:
CallExpression[callee.property.name = 'getSelectedRange']
CallExpression[callee.property.name = 'getItem']
CallExpression[callee.property.name = 'getRange']

Location the load function:
CallExpression[callee.property.name='load'][arguments.Literal.value = variableName]
*/
