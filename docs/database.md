# Database Notes

Release Radar stores append-only audit events for release state changes.
The first migration creates `audit_events` with a target index used by
support and compliance review flows.

Migration files are idempotent so a deploy job can safely retry them.

