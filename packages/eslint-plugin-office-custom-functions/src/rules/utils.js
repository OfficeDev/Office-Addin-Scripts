"use strict";
var __assign = (this && this.__assign) || function () {
    __assign = Object.assign || function(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
                t[p] = s[p];
        }
        return t;
    };
    return __assign.apply(this, arguments);
};
exports.__esModule = true;
exports.isInCustomFunction = exports.isDescendantOf = exports.isOfficeBoilerplate = exports.scopeHasLocalReference = exports.getJestFunctionArguments = exports.isDescribeEach = exports.isDescribe = exports.isTestCase = exports.getTestCallExpressionsFromDeclaredVariables = exports.isHook = exports.isFunction = exports.getNodeName = exports.TestCaseProperty = exports.DescribeProperty = exports.HookName = exports.TestCaseName = exports.DescribeAlias = exports.parseExpectCall = exports.isParsedEqualityMatcherCall = exports.EqualityMatcher = exports.ModifierName = exports.isExpectMember = exports.isExpectCall = exports.getAccessorValue = exports.isSupportedAccessor = exports.hasOnlyOneArgument = exports.getStringValue = exports.isStringNode = exports.followTypeAssertionChain = exports.createRule = void 0;
var path_1 = require("path");
var experimental_utils_1 = require("@typescript-eslint/experimental-utils");
var package_json_1 = require("../../package.json");
var REPO_URL = 'https://github.com/arttarawork/Office-Addin-Scripts';
exports.createRule = experimental_utils_1.ESLintUtils.RuleCreator(function (name) {
    var ruleName = path_1.parse(name).name;
    return REPO_URL + "/packages/eslint-plugin-office-custom-functions/blob/v" + package_json_1.version + "/docs/rules/" + ruleName + ".md";
});
var isTypeCastExpression = function (node) {
    return node.type === experimental_utils_1.AST_NODE_TYPES.TSAsExpression ||
        node.type === experimental_utils_1.AST_NODE_TYPES.TSTypeAssertion;
};
exports.followTypeAssertionChain = function (expression) {
    return isTypeCastExpression(expression)
        ? exports.followTypeAssertionChain(expression.expression)
        : expression;
};
/**
 * Checks if the given `node` is a `StringLiteral`.
 *
 * If a `value` is provided & the `node` is a `StringLiteral`,
 * the `value` will be compared to that of the `StringLiteral`.
 *
 * @param {Node} node
 * @param {V} [value]
 *
 * @return {node is StringLiteral<V>}
 *
 * @template V
 */
var isStringLiteral = function (node, value) {
    return node.type === experimental_utils_1.AST_NODE_TYPES.Literal &&
        typeof node.value === 'string' &&
        (value === undefined || node.value === value);
};
/**
 * Checks if the given `node` is a `TemplateLiteral`.
 *
 * Complex `TemplateLiteral`s are not considered specific, and so will return `false`.
 *
 * If a `value` is provided & the `node` is a `TemplateLiteral`,
 * the `value` will be compared to that of the `TemplateLiteral`.
 *
 * @param {Node} node
 * @param {V} [value]
 *
 * @return {node is TemplateLiteral<V>}
 *
 * @template V
 */
var isTemplateLiteral = function (node, value) {
    return node.type === experimental_utils_1.AST_NODE_TYPES.TemplateLiteral &&
        node.quasis.length === 1 && // bail out if not simple
        (value === undefined || node.quasis[0].value.raw === value);
};
/**
 * Checks if the given `node` is a {@link StringNode}.
 *
 * @param {Node} node
 * @param {V} [specifics]
 *
 * @return {node is StringNode}
 *
 * @template V
 */
exports.isStringNode = function (node, specifics) {
    return isStringLiteral(node, specifics) || isTemplateLiteral(node, specifics);
};
/**
 * Gets the value of the given `StringNode`.
 *
 * If the `node` is a `TemplateLiteral`, the `raw` value is used;
 * otherwise, `value` is returned instead.
 *
 * @param {StringNode<S>} node
 *
 * @return {S}
 *
 * @template S
 */
exports.getStringValue = function (node) {
    return isTemplateLiteral(node) ? node.quasis[0].value.raw : node.value;
};
/**
 * Guards that the given `call` has only one `argument`.
 *
 * @param {CallExpression} call
 *
 * @return {call is CallExpressionWithSingleArgument}
 */
exports.hasOnlyOneArgument = function (call) { return call.arguments.length === 1; };
/**
 * Checks if the given `node` is an `Identifier`.
 *
 * If a `name` is provided, & the `node` is an `Identifier`,
 * the `name` will be compared to that of the `identifier`.
 *
 * @param {Node} node
 * @param {V} [name]
 *
 * @return {node is KnownIdentifier<Name>}
 *
 * @template V
 */
var isIdentifier = function (node, name) {
    return node.type === experimental_utils_1.AST_NODE_TYPES.Identifier &&
        (name === undefined || node.name === name);
};
/**
 * Checks if the given `node` is a "supported accessor".
 *
 * This means that it's a node can be used to access properties,
 * and who's "value" can be statically determined.
 *
 * `MemberExpression` nodes most commonly contain accessors,
 * but it's possible for other nodes to contain them.
 *
 * If a `value` is provided & the `node` is an `AccessorNode`,
 * the `value` will be compared to that of the `AccessorNode`.
 *
 * Note that `value` here refers to the normalised value.
 * The property that holds the value is not always called `name`.
 *
 * @param {Node} node
 * @param {V} [value]
 *
 * @return {node is AccessorNode<V>}
 *
 * @template V
 */
exports.isSupportedAccessor = function (node, value) {
    return isIdentifier(node, value) || exports.isStringNode(node, value);
};
/**
 * Gets the value of the given `AccessorNode`,
 * account for the different node types.
 *
 * @param {AccessorNode<S>} accessor
 *
 * @return {S}
 *
 * @template S
 */
exports.getAccessorValue = function (accessor) {
    return accessor.type === experimental_utils_1.AST_NODE_TYPES.Identifier
        ? accessor.name
        : exports.getStringValue(accessor);
};
/**
 * Checks if the given `node` is a valid `ExpectCall`.
 *
 * In order to be an `ExpectCall`, the `node` must:
 *  * be a `CallExpression`,
 *  * have an accessor named 'expect',
 *  * have a `parent`.
 *
 * @param {Node} node
 *
 * @return {node is ExpectCall}
 */
exports.isExpectCall = function (node) {
    return node.type === experimental_utils_1.AST_NODE_TYPES.CallExpression &&
        exports.isSupportedAccessor(node.callee, 'expect') &&
        node.parent !== undefined;
};
exports.isExpectMember = function (node, name) {
    return node.type === experimental_utils_1.AST_NODE_TYPES.MemberExpression &&
        exports.isSupportedAccessor(node.property, name);
};
var ModifierName;
(function (ModifierName) {
    ModifierName["not"] = "not";
    ModifierName["rejects"] = "rejects";
    ModifierName["resolves"] = "resolves";
})(ModifierName = exports.ModifierName || (exports.ModifierName = {}));
var EqualityMatcher;
(function (EqualityMatcher) {
    EqualityMatcher["toBe"] = "toBe";
    EqualityMatcher["toEqual"] = "toEqual";
    EqualityMatcher["toStrictEqual"] = "toStrictEqual";
})(EqualityMatcher = exports.EqualityMatcher || (exports.EqualityMatcher = {}));
exports.isParsedEqualityMatcherCall = function (matcher, name) {
    return (name
        ? matcher.name === name
        : EqualityMatcher.hasOwnProperty(matcher.name)) &&
        matcher.arguments !== null &&
        matcher.arguments.length === 1;
};
var parseExpectMember = function (expectMember) { return ({
    name: exports.getAccessorValue(expectMember.property),
    node: expectMember
}); };
var reparseAsMatcher = function (parsedMember) { return (__assign(__assign({}, parsedMember), { 
    /**
     * The arguments being passed to this `Matcher`, if any.
     *
     * If this matcher isn't called, this will be `null`.
     */
    arguments: parsedMember.node.parent &&
        parsedMember.node.parent.type === experimental_utils_1.AST_NODE_TYPES.CallExpression
        ? parsedMember.node.parent.arguments
        : null })); };
/**
 * Re-parses the given `parsedMember` as a `ParsedExpectModifier`.
 *
 * If the given `parsedMember` does not have a `name` of a valid `Modifier`,
 * an exception will be thrown.
 *
 * @param {ParsedExpectMember<ModifierName>} parsedMember
 *
 * @return {ParsedExpectModifier}
 */
var reparseMemberAsModifier = function (parsedMember) {
    if (isSpecificMember(parsedMember, ModifierName.not)) {
        return parsedMember;
    }
    /* istanbul ignore if */
    if (!isSpecificMember(parsedMember, ModifierName.resolves) &&
        !isSpecificMember(parsedMember, ModifierName.rejects)) {
        // ts doesn't think that the ModifierName.not check is the direct inverse as the above two checks
        // todo: impossible at runtime, but can't be typed w/o negation support
        throw new Error("modifier name must be either \"" + ModifierName.resolves + "\" or \"" + ModifierName.rejects + "\" (got \"" + parsedMember.name + "\")");
    }
    var negation = parsedMember.node.parent &&
        exports.isExpectMember(parsedMember.node.parent, ModifierName.not)
        ? parsedMember.node.parent
        : undefined;
    return __assign(__assign({}, parsedMember), { negation: negation });
};
var isSpecificMember = function (member, specific) { return member.name === specific; };
/**
 * Checks if the given `ParsedExpectMember` should be re-parsed as an `ParsedExpectModifier`.
 *
 * @param {ParsedExpectMember} member
 *
 * @return {member is ParsedExpectMember<ModifierName>}
 */
var shouldBeParsedExpectModifier = function (member) {
    return ModifierName.hasOwnProperty(member.name);
};
exports.parseExpectCall = function (expect) {
    var expectation = {
        expect: expect
    };
    if (!exports.isExpectMember(expect.parent)) {
        return expectation;
    }
    var parsedMember = parseExpectMember(expect.parent);
    if (!shouldBeParsedExpectModifier(parsedMember)) {
        expectation.matcher = reparseAsMatcher(parsedMember);
        return expectation;
    }
    var modifier = (expectation.modifier = reparseMemberAsModifier(parsedMember));
    var memberNode = modifier.negation || modifier.node;
    if (!memberNode.parent || !exports.isExpectMember(memberNode.parent)) {
        return expectation;
    }
    expectation.matcher = reparseAsMatcher(parseExpectMember(memberNode.parent));
    return expectation;
};
var DescribeAlias;
(function (DescribeAlias) {
    DescribeAlias["describe"] = "describe";
    DescribeAlias["fdescribe"] = "fdescribe";
    DescribeAlias["xdescribe"] = "xdescribe";
})(DescribeAlias = exports.DescribeAlias || (exports.DescribeAlias = {}));
var TestCaseName;
(function (TestCaseName) {
    TestCaseName["fit"] = "fit";
    TestCaseName["it"] = "it";
    TestCaseName["test"] = "test";
    TestCaseName["xit"] = "xit";
    TestCaseName["xtest"] = "xtest";
})(TestCaseName = exports.TestCaseName || (exports.TestCaseName = {}));
var HookName;
(function (HookName) {
    HookName["beforeAll"] = "beforeAll";
    HookName["beforeEach"] = "beforeEach";
    HookName["afterAll"] = "afterAll";
    HookName["afterEach"] = "afterEach";
})(HookName = exports.HookName || (exports.HookName = {}));
var DescribeProperty;
(function (DescribeProperty) {
    DescribeProperty["each"] = "each";
    DescribeProperty["only"] = "only";
    DescribeProperty["skip"] = "skip";
})(DescribeProperty = exports.DescribeProperty || (exports.DescribeProperty = {}));
var TestCaseProperty;
(function (TestCaseProperty) {
    TestCaseProperty["each"] = "each";
    TestCaseProperty["concurrent"] = "concurrent";
    TestCaseProperty["only"] = "only";
    TestCaseProperty["skip"] = "skip";
    TestCaseProperty["todo"] = "todo";
})(TestCaseProperty = exports.TestCaseProperty || (exports.TestCaseProperty = {}));
var joinNames = function (a, b) {
    return a && b ? a + "." + b : null;
};
function getNodeName(node) {
    if (exports.isSupportedAccessor(node)) {
        return exports.getAccessorValue(node);
    }
    switch (node.type) {
        case experimental_utils_1.AST_NODE_TYPES.MemberExpression:
            return joinNames(getNodeName(node.object), getNodeName(node.property));
        case experimental_utils_1.AST_NODE_TYPES.NewExpression:
        case experimental_utils_1.AST_NODE_TYPES.CallExpression:
            return getNodeName(node.callee);
    }
    return null;
}
exports.getNodeName = getNodeName;
exports.isFunction = function (node) {
    return node.type === experimental_utils_1.AST_NODE_TYPES.FunctionExpression ||
        node.type === experimental_utils_1.AST_NODE_TYPES.ArrowFunctionExpression;
};
exports.isHook = function (node) {
    return node.callee.type === experimental_utils_1.AST_NODE_TYPES.Identifier &&
        HookName.hasOwnProperty(node.callee.name);
};
exports.getTestCallExpressionsFromDeclaredVariables = function (declaredVariables) {
    return declaredVariables.reduce(function (acc, _a) {
        var references = _a.references;
        return acc.concat(references
            .map(function (_a) {
            var identifier = _a.identifier;
            return identifier.parent;
        })
            .filter(function (node) {
            return !!node &&
                node.type === experimental_utils_1.AST_NODE_TYPES.CallExpression &&
                exports.isTestCase(node);
        }));
    }, []);
};
exports.isTestCase = function (node) {
    return (node.callee.type === experimental_utils_1.AST_NODE_TYPES.Identifier &&
        TestCaseName.hasOwnProperty(node.callee.name)) ||
        (node.callee.type === experimental_utils_1.AST_NODE_TYPES.MemberExpression &&
            node.callee.property.type === experimental_utils_1.AST_NODE_TYPES.Identifier &&
            TestCaseProperty.hasOwnProperty(node.callee.property.name) &&
            ((node.callee.object.type === experimental_utils_1.AST_NODE_TYPES.Identifier &&
                TestCaseName.hasOwnProperty(node.callee.object.name)) ||
                (node.callee.object.type === experimental_utils_1.AST_NODE_TYPES.MemberExpression &&
                    node.callee.object.object.type === experimental_utils_1.AST_NODE_TYPES.Identifier &&
                    TestCaseName.hasOwnProperty(node.callee.object.object.name))));
};
exports.isDescribe = function (node) {
    return (node.callee.type === experimental_utils_1.AST_NODE_TYPES.Identifier &&
        DescribeAlias.hasOwnProperty(node.callee.name)) ||
        (node.callee.type === experimental_utils_1.AST_NODE_TYPES.MemberExpression &&
            node.callee.object.type === experimental_utils_1.AST_NODE_TYPES.Identifier &&
            DescribeAlias.hasOwnProperty(node.callee.object.name) &&
            node.callee.property.type === experimental_utils_1.AST_NODE_TYPES.Identifier &&
            DescribeProperty.hasOwnProperty(node.callee.property.name));
};
/**
 * Checks if the given `describe` is a call to `describe.each`.
 *
 * @param {JestFunctionCallExpression<DescribeAlias>} node
 * @return {node is JestFunctionCallExpression<DescribeAlias, DescribeProperty.each>}
 */
exports.isDescribeEach = function (node) {
    return node.callee.type === experimental_utils_1.AST_NODE_TYPES.MemberExpression &&
        exports.isSupportedAccessor(node.callee.property, DescribeProperty.each);
};
/**
 * Gets the arguments of the given `JestFunctionCallExpression`.
 *
 * If the `node` is an `each` call, then the arguments of the actual suite
 * are returned, rather then the `each` array argument.
 *
 * @param {JestFunctionCallExpression<DescribeAlias | TestCaseName>} node
 *
 * @return {Expression[]}
 */
exports.getJestFunctionArguments = function (node) {
    return node.callee.type === experimental_utils_1.AST_NODE_TYPES.MemberExpression &&
        exports.isSupportedAccessor(node.callee.property, DescribeProperty.each) &&
        node.parent &&
        node.parent.type === experimental_utils_1.AST_NODE_TYPES.CallExpression
        ? node.parent.arguments
        : node.arguments;
};
var collectReferences = function (scope) {
    var locals = new Set();
    var unresolved = new Set();
    var currentScope = scope;
    while (currentScope !== null) {
        for (var _i = 0, _a = currentScope.variables; _i < _a.length; _i++) {
            var ref = _a[_i];
            var isReferenceDefined = ref.defs.some(function (def) {
                return def.type !== 'ImplicitGlobalVariable';
            });
            if (isReferenceDefined) {
                locals.add(ref.name);
            }
        }
        for (var _b = 0, _c = currentScope.through; _b < _c.length; _b++) {
            var ref = _c[_b];
            unresolved.add(ref.identifier.name);
        }
        currentScope = currentScope.upper;
    }
    return { locals: locals, unresolved: unresolved };
};
exports.scopeHasLocalReference = function (scope, referenceName) {
    var references = collectReferences(scope);
    return (
    // referenceName was found as a local variable or function declaration.
    references.locals.has(referenceName) ||
        // referenceName was not found as an unresolved reference,
        // meaning it is likely not an implicit global reference.
        !references.unresolved.has(referenceName));
};
exports.isOfficeBoilerplate = function (node) {
    return node.type == "CallExpression"
        && !!node.callee
        && node.callee.type == "MemberExpression"
        && node.callee.property.type == "Identifier"
        && node.callee.property.name == "run"
        && node.callee.object.type == "Identifier"
        && (node.callee.object.name == "Excel"
            || node.callee.object.name == "Word"
            || node.callee.object.name == "Powerpoint");
};
exports.isDescendantOf = function (descendantNode, ancestorNode) {
    if (descendantNode.parent === ancestorNode) {
        return true;
    }
    else {
        return descendantNode.parent ? exports.isDescendantOf(descendantNode.parent, ancestorNode) : false;
    }
};
//Requires more work
exports.isInCustomFunction = function (node, context) {
    return !!context.getSourceCode().getJSDocComment(node);
};
