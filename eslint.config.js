export default [
  { ignores: ['node_modules/**'] },
  {
    files: ['dist/**/*.js', 'eslint.config.js'],
    languageOptions: { ecmaVersion: 2022, sourceType: 'module' },
    rules: {
      'no-unused-vars': ['error', { argsIgnorePattern: '^_' }],
      'no-undef': 'off',
      'no-empty': 'error'
    }
  }
];
