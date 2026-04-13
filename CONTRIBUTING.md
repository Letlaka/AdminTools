# Contributing

Thanks for contributing to AdminTools.

## How to Contribute

1. Fork the repository and create a focused branch.
2. Make small, targeted changes.
3. Validate script behavior in your AD test environment before opening a PR.
4. Open a pull request with a clear summary, scope, and validation notes.

## Pull Request Expectations

- Keep changes scoped to one concern.
- Include updates to documentation when behavior, parameters, or outputs change.
- Preserve backward compatibility where practical.
- Do not introduce hardcoded secrets, credentials, or environment-specific sensitive data.

## Documentation Requirements

When changing `Scan-ADComputers.ps1`, update relevant docs:

- `docs/parameters.md` for parameter changes
- `docs/outputs.md` for output/report changes
- `docs/usage.md` and `docs/examples.md` for usage changes
- `CHANGELOG.md` for notable user-facing changes

## Commit Guidance

- Use clear, imperative commit messages.
- Group related edits together.
- Avoid unrelated cleanup in the same PR.

## Code and Style Notes

- Keep PowerShell changes readable and consistent with existing script style.
- Use descriptive parameter and function names.
- Prefer explicit error handling and actionable log messages.
