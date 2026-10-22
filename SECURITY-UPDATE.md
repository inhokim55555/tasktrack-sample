# TaskTrack dependency security update

Checked 2026-10-21.

The existing lockfile already audited clean. Raised the Vite minimum to 7.3.7 and added audit:security.

Validation: npm test (4 passed); npm run build passed. Full npm audit, including development dependencies: zero known vulnerabilities.

GitHub Dependabot alert listings were unavailable because alerts are disabled for this repository; this is an npm advisory check, not a claim about GitHub alert status or a complete code security audit. Existing SECURITY-REVIEW.md remains the historical source review.

Run npm ci, then npm run audit:security to repeat the dependency check.
