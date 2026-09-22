# Repository Guidelines

## Project Structure & Module Organization

This is a Rust library crate published as `justybase-spreadsheet`; the library target is named `spreadsheet` for source compatibility. Production code is in `src/`, organized by format and responsibility: XLSX/XLSB readers and writers, updaters, ZIP storage, XML helpers, and shared error/types modules. Integration tests and fixture-based compatibility checks are in `tests/` and `tests/fixtures/`. Performance comparisons are in `benches/`; runnable comparison and parity programs are in `examples/`. CI and release automation live in `.github/workflows/`.

## Build, Test, and Development Commands

- `cargo check --all-targets` — fast compile check for the library, tests, benches, and examples.
- `cargo test --all-targets` — run unit, integration, parity, and example-target tests.
- `cargo fmt --all -- --check` — verify Rust formatting.
- `cargo clippy --all-targets --all-features -- -D warnings` — run lint checks with warnings treated as errors.
- `cargo bench --bench read_comparison -- --quick` — run the XLSX/XLSB reader comparison benchmark and compatibility assertions.
- `cargo package --allow-dirty --list` — verify package contents before release. Publishing is performed by GitHub Actions from a matching `vX.Y.Z` tag.

Run the complete local gate with `cargo test --all-targets && cargo fmt --all -- --check && cargo clippy --all-targets --all-features -- -D warnings`.

## Coding Style & Naming Conventions

Use standard `rustfmt` formatting, four-space indentation, `snake_case` for functions/modules, `UpperCamelCase` for types, and `SCREAMING_SNAKE_CASE` for constants. Prefer small format-specific helpers and propagate errors with the project’s `SpreadsheetResult`/`SpreadsheetError` types. Preserve deterministic ZIP timestamps and output ordering when changing writers.

## Testing Guidelines

Add focused unit tests beside the relevant module for parsing or serialization behavior. Add integration tests under `tests/` when validating cross-module, cross-runtime, or fixture compatibility. Keep parity assertions meaningful and run `cargo test --all-targets` before submitting changes.

## Commit & Pull Request Guidelines

Use concise, imperative Conventional Commit-style messages, for example `chore: release version 0.1.1` or `fix: preserve worksheet ordering`. Pull requests should explain behavior changes, list validation commands run, and call out compatibility or performance effects. Release changes must update `Cargo.toml`/`Cargo.lock`; publish by pushing the matching version tag so the release workflow can validate and publish through GitHub Actions.
