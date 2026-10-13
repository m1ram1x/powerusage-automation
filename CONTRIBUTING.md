# Contributing

Use synthetic CSV fixtures and mocked Telegram requests. Keep private measurements and credentials outside commits.

For a change:

1. Explain the input behavior being changed and the expected report/alert.
2. Preserve the supported CSV format, kW interpretation and template layout unless a behavior change is intentional and documented.
3. Run `python -m unittest discover -s tests -v` in your prepared Python environment.
4. For container/launcher changes, run the Docker demo and test the relevant operating system.
5. Update README/help text when command flags or output behavior change.

Report parsing bugs with a small anonymized CSV, the command used and the exact error. Never attach `telegram.env` or a real token.
