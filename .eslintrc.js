require('@rushstack/eslint-config/patch/modern-module-resolution');
module.exports = {
  extends: ['@microsoft/eslint-config-spfx/lib/profiles/default'],
  parserOptions: { tsconfigRootDir: __dirname },
  rules: {
    // Purely stylistic; we keep explicit annotations for readability.
    '@typescript-eslint/no-inferrable-types': 'off',
    // The ExcelJS / Chart.js interop in the export code is heavily dynamic;
    // typing it fully would add noise without real safety.
    '@typescript-eslint/no-explicit-any': 'off'
  },
  overrides: [
    {
      files: ['*.ts', '*.tsx'],
      parser: '@typescript-eslint/parser',
      parserOptions: { project: './tsconfig.json' }
    }
  ]
};
