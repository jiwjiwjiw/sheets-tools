# sheets-tools

Helpers shared by Google Apps Script projects bound to Google Sheets:

- `Validation` / `ValidationHandler`: keep data-validation dropdowns of a
  sheet in sync with the values of another range plus fixed entries.
- `helpers`: row filters, Drive id extraction from sharing links, query
  escaping, range intersection.
- `EmailTemplate`: build an HTML email from a template sheet with rich text.

The module is embedded as a git submodule (`src/sheets-tools`) by
invoice-manager and voluntary-manager; it has no build of its own. Commit and
push changes here first, then record the new commit in each consumer.

## Checks

```sh
yarn install
yarn typecheck   # strictest TypeScript options of the consumers
yarn lint        # ESLint, same configuration as the consumers
```

Both run in CI on every push.
