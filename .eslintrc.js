module.exports = {
  root: true,
  parser: '@typescript-eslint/parser',
  plugins: ['@typescript-eslint'],
  extends: ['eslint:recommended', 'plugin:@typescript-eslint/recommended'],
  env: { node: true },
  rules: {
    '@typescript-eslint/no-unused-vars': ['warn', { args: 'none' }],
    'comma-dangle': [
      'error',
      {
        arrays: 'always-multiline',
        objects: 'always-multiline',
        imports: 'always-multiline',
        exports: 'always-multiline',
        functions: 'always-multiline',
      },
    ],
    'func-call-spacing': ['error', 'never'],
    indent: ['error', 2],
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
    quotes: [2, 'single', { avoidEscape: true }],
    semi: ['error', 'always'],
    'space-in-parens': ['error', 'never'],
  },
};
