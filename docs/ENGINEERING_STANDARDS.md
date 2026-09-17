# Engineering Standards

This automation is treated as a production reporting pipeline. Changes should preserve data integrity before convenience.

## Safety

- Never overwrite source workbooks; write a new output artifact.
- Keep real/customer workbooks out of Git history.
- Treat UPC as the stable join key and validate unexpected stores/categories.
- Reject invalid configuration and non-finite numeric values early.

## Testing

- Keep parsing and validation logic deterministic and independently testable.
- Add regression coverage for every repaired business rule.
- Prefer pure helpers for transformations that do not require Excel COM.
- Excel COM behavior should be isolated from the data layer.

## Reconciliation

Every monthly run should make it possible to explain new, missing, unmapped, and unknown records. Grand totals must reconcile to the underlying units and dollar totals before delivery.

## Change management

Pull requests must contain genuine engineering work. Repository history, LOC, and PR counts must never be fabricated or backdated.
