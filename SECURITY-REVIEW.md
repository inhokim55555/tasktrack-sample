# Security Review: Sample-Repositories

## Scope

Security and publication-readiness review of BYOU Intel, Trackify and TaskTrack local sample source.

- Scan mode: repository
- Target kind: directory_snapshot
- Target ID: target_sha256_934576398b3928b7b746aeb2fa60b404816f6dca44112f2818400e6562cc2000
- Snapshot digest: codex-security-snapshot/v1:sha256:6225b5342eab2d2d124e81469a8fa779832692ca612c0091f7bd4e3c97e06daf
- Inventory strategy: directory
- Included paths: .
- Excluded paths: none
- Artifacts reviewed: artifacts/validation.md

Limitations and exclusions:
- Not an assurance of absence of vulnerabilities. Third-party dependency implementation and binary forensics excluded.
- Trackify retains nine development-only dependency audit entries (seven high, two moderate); production dependency audits report zero across all samples.
- Source changed after the initial scan snapshot during remediation; this report does not cryptographically certify later publication commits.
- Excluded \*\*/node_modules/\*\*, generated builds: node_modules implementation and generated build internals; dependency advisories checked separately.
- Excluded binary assets and host OS: Binary forensic examination and operating-system network sandbox certification.
- Excluded concurrent collection moves: Stress testing simultaneous collection moves.

### Scan Summary

| Field | Value |
| --- | --- |
| Scan outcome | completed |
| Reportable findings | 0 |
| Severity mix | none |
| Confidence mix | none |
| Coverage | partial |
| Validation mode | Full authored-source review by independent baseline/architecture reviewers plus coordinator remediation and runtime regression checks. |

Canonical artifacts: `scan-manifest.json`, `findings.json`, and `coverage.json`. This report is a deterministic projection of those files.

## Threat Model

Three local sample applications: Intel Node/React with file datasets and local PGlite P&L; Trackify Next.js/Prisma with local PGlite; TaskTrack browser-local task boards. sample/start.mjs now supplies loopback ports and isolated child environment.

### Assets

- Publication source/seed/assets must exclude customer data and credentials.
- Intel sample-state/intel file datasets and sample-state/database P&L (byou-intel/sample/start.mjs).
- Trackify fictional Prisma state (trackify/scripts/portfolio-seed.mjs); TaskTrack browser localStorage (tasktrack/src/taskManagerApp.js).

### Trust Boundaries

- Public repository readers can inspect all source; runtime isolation cannot protect published PII.
- Browser to loopback proxy requires exact Host and mutation Origin (\*/sample/start.mjs); direct Trackify backend also checks Host+Origin (trackify/src/proxy.ts).
- Child environment is allowlisted; mock providers and fixed database targets isolate sample data (trackify/src/lib/db.ts, datasource/index.ts).
- TaskTrack Supabase consumer is null regardless of build variables (tasktrack/src/supabaseClient.js).

### Attacker Capabilities

- Remote websites can attempt cross-origin or DNS-rebinding requests; local operators can modify source but already have that authority.
- Visitors can intentionally edit fictional local records; this is not production authentication.

### Security Objectives

- Publish reviewed fictional data without source history or credentials.
- Avoid production services and sensitive host files.
- Bound request parsing and preserve safe rendering.

### Assumptions

- This is a local fictional-data demonstration, not production authentication.
- Source was remediated during review; initial directory snapshot predates fixes. Final build/test receipts describe revised source.
- OS isolation, third-party dependency internals and binary forensics are outside this review.

## Findings

### No findings

No reportable findings survived the canonical discovery, validation, and reportability gates.

## Reviewed Surfaces

| Surface | Risk Area | Outcome | Notes |
| --- | --- | --- | --- |
| track-host | not recorded | Rejected | Exact loopback Host enforced at frontend and raw backend; hostile Host returned 403 on five ports. |
| track-body | not recorded | Rejected | Actual streamed bytes are bounded; missing Content-Length oversized regression passes. |
| intel-offline | not recorded | Rejected | Production clients/internal research modules removed; fixed local data/database targets; external socket/fetch tests pass; restrictive CSP. |
| track-cycle | not recorded | Rejected | Visited-set traversal and transaction advisory lock serialize reparent validation/update. Source reviewed; concurrent stress test excluded. |
| track-csv | not recorded | Rejected | Export uses formula-neutralizing helper; whitespace-prefixed formula regression passes. |
| intel-private-fixtures | not recorded | Rejected | Owner approval ledger emptied, real product aliases and inherited demographic exports removed; only generated operational seed retained. |
| task-sanitize | not recorded | Rejected | Descendants cleaned before unwrapping; nested dangerous element/attribute regression passes. |
| task-labels | not recorded | Rejected | Editable labels/icons escaped before HTML insertion; source reviewed. |
| track-gmail-get | not recorded | Rejected | Simulated mailbox mutation requires same-origin Sec-Fetch-Site. No real OAuth or tokens. |

## Open Questions And Follow Up

- Before adapting these samples for hosting, restore real authentication and repeat a deployment-specific review.
- Trackify has nine development-only npm audit entries (seven high, two moderate) through Tailwind/ESLint: braces stack exhaustion and postcss-selector-parser CPU complexity. Runtime audit is zero; no normal request reaches build pattern inputs. Major build-tool migration remains deferred.
  - Follow-up prompt: Review deferred unit track-build-advisories and close its stated proof gap.
