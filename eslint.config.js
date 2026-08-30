const commonGlobals = {
    Buffer: 'readonly',
    console: 'readonly',
    DOMParser: 'readonly',
    Document: 'readonly',
    Element: 'readonly',
    fetch: 'readonly',
    global: 'readonly',
    globalThis: 'readonly',
    Node: 'readonly',
    process: 'readonly',
    setTimeout: 'readonly',
    clearTimeout: 'readonly',
    TextDecoder: 'readonly',
    TextEncoder: 'readonly',
    URL: 'readonly',
    XMLSerializer: 'readonly'
};

const productionFiles = [
    'index.js',
    'adapters/**/*.js',
    'core/**/*.js',
    'engine/**/*.js',
    'pipeline/**/*.js',
    'services/**/*.js',
    'orchestration/**/*.js'
];

const wordElementFactoryRestrictions = [
    {
        selector: "CallExpression[callee.type='MemberExpression'][callee.property.name='createElement'][arguments.0.value=/^w:/]",
        message: 'Create WordprocessingML elements with createWordElement().'
    },
    {
        selector: "CallExpression[callee.type='MemberExpression'][callee.property.name='createElementNS'][arguments.0.name='NS_W']",
        message: 'Create WordprocessingML elements with createWordElement().'
    },
    {
        selector: "CallExpression[callee.type='MemberExpression'][callee.property.name='createElementNS'][arguments.0.value='http://schemas.openxmlformats.org/wordprocessingml/2006/main']",
        message: 'Create WordprocessingML elements with createWordElement().'
    }
];

export default [
    {
        ignores: ['coverage/**', 'dist/**', 'node_modules/**', 'tmp/**']
    },
    {
        files: ['**/*.js', '**/*.mjs'],
        languageOptions: {
            ecmaVersion: 'latest',
            sourceType: 'module',
            globals: commonGlobals
        },
        rules: {
            'no-empty': ['error', { allowEmptyCatch: false }],
            'no-undef': 'error',
            'no-unused-vars': ['error', {
                argsIgnorePattern: '^_',
                caughtErrors: 'none',
                varsIgnorePattern: '^_'
            }],
            'require-atomic-updates': 'error'
        }
    },
    {
        files: productionFiles,
        rules: {
            'no-restricted-syntax': [
                'error',
                ...wordElementFactoryRestrictions,
                {
                    selector: "CallExpression[callee.type='MemberExpression'][callee.property.name='parseFromString']",
                    message: 'Parse XML through adapters/xml-adapter.js.'
                }
            ]
        }
    },
    {
        files: ['core/word-xml.js'],
        rules: {
            'no-restricted-syntax': ['error', {
                selector: "CallExpression[callee.type='MemberExpression'][callee.property.name='parseFromString']",
                message: 'Parse XML through adapters/xml-adapter.js.'
            }]
        }
    },
    {
        files: ['adapters/xml-adapter.js'],
        rules: {
            'no-restricted-syntax': ['error', ...wordElementFactoryRestrictions]
        }
    }
];
