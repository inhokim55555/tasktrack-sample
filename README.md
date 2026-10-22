# TaskTrack Sample Repo

A standalone portfolio sample adapted from application source at revision `617a0e28fdde3ca642c73c4ae59d682c7ea622cd`. This repository starts with a fresh history and contains no production database or credentials.

Task boards and documents stored in your browser. Assistant and meeting analysis are demonstrations. Cloud meeting saving, uploads, mail delivery and external account connections are disabled.

## Run locally

1. Install Node.js 24, then download and extract this repository.
2. Open a terminal in this folder and run `npm ci`. This installation needs internet access.
3. Run `npm run build`.
4. Run `npm start` (or double-click `START SAMPLE.cmd` after the first build), then open http://127.0.0.1:4296/. Press Ctrl+C in the terminal to stop.

Run `npm test` for the included security regression checks.

## Sample boundaries

Use only fictional information. The app intentionally has a local sample identity instead of production sign-in. Do not deploy it publicly as a hosted service or enter real customer information. The launcher binds to your computer only, uses fixed local database addresses and rejects external service calls. These controls are application safeguards, not an operating-system sandbox.

Local database state lives in ignored `sample-state/`; TaskTrack browser edits remain in browser storage. The repository excludes dependencies, generated builds, runtime state, environment files and original Git history.

## Security review

See [SECURITY-REVIEW.md](SECURITY-REVIEW.md) for review scope, fixes, tests and limitations. A review does not guarantee absence of vulnerabilities.
