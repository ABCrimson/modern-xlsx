## Summary

<!-- Brief description of what this PR does -->

## Changes

<!-- Bullet list of specific changes -->

-

## Testing

<!-- How was this tested? -->

Tick the commands you ran. CI runs each of them (`.github/workflows/ci.yml`), along with further jobs (fuzz formatting, the `parallel` feature, rustdoc, browser tests, WASM builds, examples); see the job table in CONTRIBUTING.md.

- [ ] `cargo fmt --all --check`
- [ ] `cargo clippy --workspace --all-targets --locked -- -D warnings`
- [ ] `cargo clippy --workspace --all-targets --no-default-features --locked -- -D warnings`
- [ ] `cargo test --workspace --locked`
- [ ] `cargo test --workspace --no-default-features --locked`
- [ ] `pnpm -C packages/modern-xlsx typecheck`
- [ ] `pnpm -C packages/modern-xlsx typecheck:tests`
- [ ] `pnpm lint` (Biome over the whole repository: docs, examples, scripts)
- [ ] `pnpm -C packages/modern-xlsx lint`
- [ ] `pnpm -C packages/modern-xlsx build`
- [ ] `pnpm -C packages/modern-xlsx test`

## Checklist

- [ ] Golden diffs regenerated with `UPDATE_GOLDEN=1` were reviewed
- [ ] A user-visible change has an entry under Unreleased in `packages/modern-xlsx/CHANGELOG.md`
- [ ] Docs and comments are updated to match the code

## Related Issues

<!-- Link related issues: Fixes #123, Relates to #456 -->
