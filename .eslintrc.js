module.exports = {
  root: true,
  parser: '@typescript-eslint/parser',
  plugins: ['@typescript-eslint'],
  extends: ['eslint:recommended', 'plugin:@typescript-eslint/recommended'],
  env: { node: true },
  rules: {
    quotes: [2, 'single', { avoidEscape: true }],
    'no-unused-vars': 'off',
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
  },
};
