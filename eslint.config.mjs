// Keep in sync with the eslint.config.mjs of the projects that embed this
// module (invoice-manager, voluntary-manager): they lint these files with
// their own copy of this configuration.
import eslint from '@eslint/js'
import tseslint from 'typescript-eslint'

const nodeGlobals = {
  require: 'readonly',
  module: 'writable',
  process: 'readonly',
  __dirname: 'readonly'
}

export default tseslint.config(
  { ignores: ['dist/**', 'node_modules/**'] },
  eslint.configs.recommended,
  {
    files: ['**/*.ts'],
    extends: [...tseslint.configs.recommended],
    rules: {
      // Apps Script code throws plain strings on purpose: they are shown to the
      // user as-is in a dialog.
      '@typescript-eslint/only-throw-error': 'off',
      // Spreadsheet cell values are untyped by nature.
      '@typescript-eslint/no-explicit-any': 'off'
    }
  },
  {
    files: ['**/*.js'],
    languageOptions: { sourceType: 'commonjs', globals: nodeGlobals }
  },
  {
    files: ['**/*.mjs'],
    languageOptions: { sourceType: 'module', globals: nodeGlobals }
  }
)
