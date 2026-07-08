# Release Radar

Release Radar is a tiny public sample app for the Kinetik Triage live example.
It exposes release-readiness data for a fictional SaaS team without using
Kinetik source code or secrets.

The repository intentionally has ordinary product history: feature work,
a schema migration, an authentication change, tests, a tag, and a GitHub
Release. Kinetik uses that public history as verifiable evidence for its
example analysis.

## Run

```bash
npm test
npm start
```

The server listens on `PORT` or `3000`.

## Endpoints

- `GET /health` returns a liveness check.
- `GET /api/releases` returns release records.
- `GET /api/readiness` returns scored release-readiness summaries.
