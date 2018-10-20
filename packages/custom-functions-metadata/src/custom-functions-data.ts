#!/usr/bin/env node

// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

namespace CustomFunctionMetadata {
    export interface Metadata {
        functions: Function[];
    }
    
    export interface Function {
        name: string;
        id: string;
        helpUrl: string;
        description: string;
        parameters: FunctionParameter[];
        result: FunctionResult;
        options: FunctionOptions;
    }
    
    export interface FunctionOptions {
        volatile: boolean;
        stream: boolean;
        cancelable: boolean;
    }
    
    export interface FunctionParameter {
        name: string;
        description?: string;
        type: string;
        dimensionality: string;
        optional: boolean;
    }
    
    export interface FunctionResult {
        type: string;
        dimensionality: string;
    }
}