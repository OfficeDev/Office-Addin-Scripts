import {
  getPropertyNameInLoad,
  findPropertiesRead,
  findOfficeApiReferences,
  OfficeApiReference,
} from "../utils";

export = {
  name: "call-sync-after-load",
  meta: {
    type: <"problem" | "suggestion" | "layout">"suggestion",
    messages: {
      callSyncAfterLoad:
        "Call context.sync() after calling load '{{name}}' on '{{loadValue}}' and before reading property",
    },
    docs: {
      description:
        "Always call load on an object followed by a context.sync() before reading it or one of its properties.",
      category: <
        "Best Practices" | "Stylistic Issues" | "Variables" | "Possible Errors"
      >"Best Practices",
      recommended: <false | "error" | "warn">false,
      url: "https://docs.microsoft.com/en-us/office/dev/add-ins/develop/application-specific-api-model?view=powerpoint-js-1.1#load",
    },
    schema: [],
  },
  create: function (context: any) {
    type VariableProperty = {
      variable: string;
      property: string;
    };

    class VariablePropertySet extends Set {
      add(variableProperty: VariableProperty) {
        return super.add(JSON.stringify(variableProperty));
      }
      has(variableProperty: VariableProperty) {
        return super.has(JSON.stringify(variableProperty));
      }
    }

    let apiReferences: OfficeApiReference[] = [];

    function findLoadBeforeSync(): void {
      const needSync: VariablePropertySet = new VariablePropertySet();
      const wasLoaded: VariablePropertySet = new VariablePropertySet();
      let hasSync = false;

      apiReferences.forEach((apiReference) => {
        const operation = apiReference.operation;
        const reference = apiReference.reference;
        const variable = reference.resolved;

        if (operation === "Load" && variable) {
          const propertyName: string = getPropertyNameInLoad(
            reference.identifier.parent
          );
          needSync.add({ variable: variable.name, property: propertyName });
          wasLoaded.add({ variable: variable.name, property: propertyName });
        }

        if (operation === "Sync") {
          hasSync = true;
          needSync.clear();
        }

        if (operation === "Read" && variable) {
          const propertyName: string = findPropertiesRead(
            reference.identifier.parent
          );
          const variableProperty: VariableProperty = { variable: variable.name, property: propertyName };
          if (
            needSync.has(variableProperty) 
            && wasLoaded.has(variableProperty) 
            && hasSync
          ) {
            const node = reference.identifier;
            context.report({
              node: node,
              messageId: "callSyncAfterLoad",
              data: { name: node.name, loadValue: propertyName },
            });
          }
        }
      });
    }

    return {
      Program() {
        apiReferences = findOfficeApiReferences(context.getScope());
        apiReferences.sort((left, right) => {
          return (
            left.reference.identifier.range[1] -
            right.reference.identifier.range[1]
          );
        });
        findLoadBeforeSync();
      },
    };
  },
};
