module.exports = {
  root: true,
  env: {
    browser: true,
    es2021: true,
    node: true
  },
  parserOptions: {
    ecmaVersion: 2021,
    sourceType: 'script'
  },
  extends: ['eslint:recommended'],
  overrides: [
    {
      files: ['Code.gs'],
      rules: {
        'no-unused-vars': 'off'
      }
    },
    {
      files: ['*.html'],
      parser: '@html-eslint/parser',
      plugins: ['@html-eslint'],
      rules: {
        '@html-eslint/indent': 'off',
        '@html-eslint/no-extra-spacing-attrs': 'off'
      }
    },
    {
      files: ['tests/**/*.js'],
      env: {
        node: true,
        es2021: true
      },
      parserOptions: {
        sourceType: 'module'
      }
    }
  ],
  globals: {
    CacheService: 'readonly',
    HtmlService: 'readonly',
    LockService: 'readonly',
    Logger: 'readonly',
    MailApp: 'readonly',
    PropertiesService: 'readonly',
    Session: 'readonly',
    SpreadsheetApp: 'readonly',
    Utilities: 'readonly',
    include: 'readonly'
  },
  ignorePatterns: ['node_modules/']
};
