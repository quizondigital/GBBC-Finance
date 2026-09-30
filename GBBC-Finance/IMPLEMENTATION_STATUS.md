# GBBC portal implementation status

This project is a working local Flask implementation of the GBBC Finance and Giving Portal. Its green-and-cream dashboard, sidebar, cards, tables, and logo follow the supplied proposal PDF. It is separate from the older Next.js prototype in `../gbbc-finance-portal` because the requirements specify Flask and MySQL.

## Implemented

- Sign-in, logout, salted password hashes, idle-session expiration, login throttling, CSRF checks, and audit records.
- Four defined roles, server-side route enforcement, suppressed navigation and detail for unauthorized users, and an administrator-visible role-permissions matrix.
- User creation and editing with required first name, last name, email, phone, role, and password; optional username; password reset, deactivation, and last-Super-Admin protection.
- People directory, create/edit forms, search, CSV template, and validated CSV import.
- Contribution entry, correction, voiding, reconciliation status, fiscal-year constraints, decimal currency storage, and fund association.
- Aggregate and detailed reports, monthly trends, date and year filters, contributor threshold, Excel export, and individual PDF acknowledgment generation.
- Fiscal-year creation and open/close controls, church identity settings, feedback submission/review, and audit-history page.
- MySQL configuration through `DATABASE_URL`, PythonAnywhere WSGI entry point, HTTPS-capable cookie configuration, and responsive desktop/mobile layout.

## Before live use

- Deploy to the church's PythonAnywhere account, connect its MySQL database, configure DreamHost DNS and a valid HTTPS certificate, and run production acceptance checks. These external accounts were not available in this workspace.
- Confirm the church's legal name, mailing address, and approved acknowledgment language. The PDF statement uses provisional wording until approved.
- Add and verify schema migrations, backup/restore automation, retention and export-cleanup procedures, monitoring, browser/accessibility testing, and a full security review.
- Complete remaining detailed requirements such as richer person fields and CSV reconciliation/duplicate preview, configurable funds and fiscal close checklist, report refinements, and production integration settings. The current release is a functional implementation, not a claim that every Must item in the 21-page specification has passed acceptance.
