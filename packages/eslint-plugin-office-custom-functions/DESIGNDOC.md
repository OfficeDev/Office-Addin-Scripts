# Shared App Linter Design Doc (Rough)

This linter is built in eslint. It's purpose is to throw errors whenever a user tries to run Office Api calls, so there's less resource waste.

There are three general phases of the linting, with the third phase being the most important. The pseudocode below is based off the prototype code written within Office-Script-Linter.

## Phase 1: Look for Excel.run()-type Call statement

```javascript

    public function lint(): void {
        rootNode.forEachChild(traverseNode);
    }

    function traverseNode(node) {
        if (
            isCallExpression(node)
            && isPropertyAccessExpression(node.expression)
            && isIdentifier(node.expression.expression)
        ) {
            if (isOfficeBoilerPlate(node.expression)) { //checks if node is in Excel.run() type format
                node.forEachChild(findBlockNode);
                return;
            }
        }
        node.forEachChild(traverseNode);
    }

```

## Phase 2: Parse through code block

```javascript

    function findBlockNode(node) {
        if (isBlock(node)) {
            node.forEachChild(lintNode);
        } else {
            node.forEachChild(findBlockNode);
        }
    }

```

## Phase 3: Evaluate if funct is Office API call or not

```javascript

    function lintNode(node) {
        if (isIdentifier(node)) {
            if (validateApiMethod(node)) {
                createError(
                    LinterErrorCode.NoReadApiCall,
                    {
                        node.startChar,
                        node.startLine,
                        node.endChar,
                        node.endLine,
                    },
                    LinterErrorSeverity.Error
                );
                this.counter.incrementCount(
                    LinterErrorCode.NoReadApiCall
                );
            }
        }
        if (
            isForStatement(node) ||
            isWhileStatement(node)
        ) {
            node.forEachChild(findBlockNode);
        } else {
            node.forEachChild(lintNode);
        }
    }

```

### Phase 3a: validateApiMethod()

This is where there could be difficulties, due to the difference between the eslint ast format and the typescript ast format. Typing information is easier to get from the Typescript AST, so converting this to an eslint format can be tricky. Pictured below is the typescript implementation of this in the old prototype

```typescript

export function validateWriteMethod(
    apiReferences: OfficeScriptJson,
    namespace?: string,
    objectName?: string,
    methodName?: string
) {
    if (!objectName || !methodName) {
        return false;
    }

    const { namespaces } = apiReferences;

    const foundNamespace = Object.keys(namespaces).find(ns => ns === namespace);
    if (!foundNamespace) {
        return false;
    }

    const { objects } = namespaces[foundNamespace];
    if (!objects[objectName]) {
        return false;
    }

    if (objects[objectName].methods[methodName]) {
        if (!objects[objectName].methods[methodName].isRead) {
            return true;
        }
    }

    return false;
}

```