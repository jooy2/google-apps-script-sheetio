module.exports = {
  env: {
    browser: true,
    es2021: true,
  },
  extends: [
    'airbnb-base',
  ],
  parserOptions: {
    ecmaVersion: 12,
    sourceType: 'module',
  },
  rules: {
    'no-undef': 0,
    'no-throw-literal': 0,
    'no-unused-expressions': 1,
    'no-unused-vars': 1,
    'space-before-function-paren': ['error', 'always'],
  },
};
