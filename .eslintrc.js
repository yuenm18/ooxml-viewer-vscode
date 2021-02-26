module.exports = {
  root: true,
  parser: '@typescript-eslint/parser',
  plugins: ['@typescript-eslint'],
  extends: ['eslint:recommended', 'plugin:@typescript-eslint/recommended'],
  env: { node: true },
  rules: {
    '@typescript-eslint/no-unused-vars': ['warn', { args: 'none' }],
    camelcase: ['warn', { properties: 'always' }],
    'comma-dangle': [
      'error',
      {
        arrays: 'always-multiline',
        objects: 'always-multiline',
        imports: 'never',
        exports: 'always-multiline',
        functions: 'always-multiline',
      },
    ],
    'func-call-spacing': ['error', 'never'],
    indent: ['error', 2, { SwitchCase: 1 }],
    'max-len': [
      'warn',
      {
        code: 140,
        tabWidth: 2,
        ignoreComments: true,
        ignoreUrls: true,
        // ignoreStrings: true,
      },
    ],
    'no-console': ['warn', { allow: ['warn', 'error'] }],
    'no-unused-vars': 'off',
    quotes: ['error', 'single', { avoidEscape: true }],
    semi: ['error', 'always'],
    'space-in-parens': ['error', 'never'],
    'newline-per-chained-call': ['error', { ignoreChainWithDepth: 3 }],
  },
};
