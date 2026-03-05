Below is a **requirements set for a web-based platform** that supports running the CPMS evaluation test suite you built (categories → tests → scoring → evidence → reporting), without recreating the over-detailed RFP text. It’s written so you can hand it to a dev team or vendor as a product spec.

---

## 1) Purpose and scope

### Objective

Provide a web-based platform to **configure, run, score, and report** CPMS vendor evaluations using a structured test suite (functional + non-functional), aligned to **categories**, with **evidence capture**, **auditability**, and **weighted scoring**.

### In scope

* Test suite authoring/import, execution workflows, scoring, evidence, reporting, collaboration, and versioning.
* Multi-vendor evaluation across the same suite, including comparison views.

### Out of scope (unless explicitly added)

* Running live charge-point integrations (OCPP simulators, etc.)
* Contracting, procurement approvals, or vendor onboarding portals (can be integrated later)

---

## 2) Users, roles, and permissions

### Roles (minimum)

* **Admin**: org settings, user mgmt, security, templates, global scoring rules.
* **Evaluation Manager**: creates projects, assigns evaluators, finalizes results, publishes reports.
* **Evaluator**: executes assigned tests, records scores, uploads evidence, adds notes.
* **Reviewer/Approver**: reviews evidence, requests changes, approves final scoring.
* **Vendor User (optional)**: can view assigned tests, submit evidence, respond to clarifications (no ability to change scores unless explicitly allowed).

### Permissions model

* Role-based access control with the ability to set:

  * **Project-level roles**
  * **Category-level ownership**
  * **Test-level assignments**
* Full audit trail for changes to: suite, tests, scoring, weights, and approvals.

---

## 3) Core data model (conceptual)

### Entities

* **Organization / Workspace**
* **Users**
* **Evaluation Project**
* **Vendor (and vendor contacts)**
* **Test Suite**
* **Suite Version** (immutable once published)
* **Category** (hierarchical: Category → Subcategory optional)
* **Test Case**
* **Test Run** (project + vendor + suite version)
* **Test Result** (test case instance with score, status)
* **Evidence Item** (file/link/text, metadata)
* **Comment / Discussion Thread**
* **Scorecard** (computed aggregates)
* **Approval Record** (sign-off steps)

### Key relationships

* A project can evaluate **multiple vendors**
* Each vendor has a **test run** against a **specific suite version**
* Each test case produces a **test result** per vendor/run

---

## 4) Functional requirements

### 4.1 Test suite management

* Create/edit a suite with:

  * Categories (aligned to your RFP structure)
  * Test cases with: tier (Core/Extended), priority, weight, method (Demo/Docs/POC), pass criteria, evidence required, traceability fields.
* **Import** from Excel (your CPMS suite):

  * Map columns → platform fields
  * Validate required fields
  * Report import errors/warnings
* **Export** back to Excel/PDF with the same structure
* **Versioning**:

  * Draft versions editable
  * Published versions immutable
  * “Clone suite” creates a new draft version
* Library features:

  * Templates (baseline CPMS suite)
  * Reusable test case blocks per category

### 4.2 Project setup and vendor onboarding

* Create evaluation project with:

  * Scope (which categories/tests included)
  * Scoring scheme (see 4.4)
  * Timeline & milestones (optional)
* Add vendors and assign:

  * Vendor users (optional)
  * NDAs / access notes (metadata only)
* Define evaluation method per vendor:

  * Demo-only, sandbox/POC, pilot (used to enforce evidence requirements)

### 4.3 Test execution workflow

* Assign tests to evaluators by category/test
* Test states (minimum):

  * **Not Started**
  * **In Progress**
  * **Blocked** (with reason)
  * **Submitted for Review**
  * **Approved**
* For each test result, support:

  * Score selection (dropdown scale)
  * Pass/Fail (optional separate field)
  * Notes (rich text)
  * Evidence attachments (files) + evidence links + screenshots
  * Vendor response/clarification thread (optional)
* Bulk actions:

  * Bulk set “method”, “owner”, “status”, “tier filter”, “N/A”
* Time-stamped activity log per test result

### 4.4 Scoring, weighting, and normalization

* Configurable scoring scale:

  * Example: OOB=5, Configurable=4, Custom=2, Roadmap=1, Not supported=0, N/A blank
* Weighting at:

  * Test level (required)
  * Optional category weight multiplier
* Computations:

  * Weighted score per test and category
  * Overall vendor score
  * Coverage metrics: % of core tests executed, % evidence-complete
* Rules:

  * “Final score locked only after approval”
  * Ability to “freeze scoring scheme” per project
* Support “gating requirements”:

  * Mark tests as **Mandatory**; failing them flags vendor as non-compliant regardless of total score

### 4.5 Evidence management

* Evidence types:

  * Upload file (pdf, image, doc, zip)
  * External link (URL)
  * Text note / transcript
* Evidence metadata:

  * Who uploaded, when, for which test, vendor, and suite version
  * Evidence status: pending / accepted / rejected
* Evidence requirements:

  * For tests marked “Evidence Required”, block submission without evidence
* Retention & access controls for evidence per project

### 4.6 Reporting and dashboards

* Dashboards:

  * Progress by vendor/category
  * Blockers and overdue items
  * Evidence completeness
* Reports (exportable PDF + web view):

  * Vendor scorecard
  * Category breakdown
  * Core-only summary
  * Mandatory/gating requirements status
  * Top gaps (lowest-scoring categories/tests)
* Comparison views:

  * Compare vendors across categories
  * Compare vendors on a subset of core tests
* Traceability report:

  * Show mapping from original requirement IDs to test cases (if included)

### 4.7 Collaboration and review

* Comments per test (threaded)
* @mentions and notifications
* Review workflow:

  * Reviewer can request changes
  * Approver can lock results
* Change tracking:

  * show what changed, by whom, and when (suite/test definitions and results)

### 4.8 Admin and configuration

* Org settings:

  * Default scoring scale
  * Default evidence policy
  * SSO config (see NFR)
* User management:

  * Invite, deactivate, role assignment
* Field customization (lightweight):

  * Add custom tags (e.g., “OCPP”, “ISO 15118”, “Energy Mgmt”)
* Data export:

  * Full project export (results + evidence index + logs)

---

## 5) Non-functional requirements (NFR)

### Security & compliance

* **SSO** (SAML/OIDC) + optional MFA
* Encryption:

  * In transit (TLS 1.2+)
  * At rest for DB and stored evidence
* RBAC + least privilege
* Audit logs for all sensitive actions
* Vulnerability management + secure SDLC expectations
* Compliance targets depending on procurement: SOC 2 Type II / ISO27001 (optional requirement)

### Availability & performance

* Availability target: e.g., **99.9%** monthly
* Performance:

  * Dashboard loads < 3s for typical projects (e.g., 200–500 tests × 5 vendors)
  * Bulk updates for 100+ test results under 10s
* Scalability:

  * Support multiple concurrent evaluation projects
* Backups:

  * Daily backups + point-in-time recovery (if cloud DB supports)

### Data & privacy

* GDPR alignment:

  * Data minimization
  * Right to delete user accounts
  * EU data residency option (if required)
* Retention policies configurable per org/project

### Usability

* Fully web-based, responsive UI
* Accessibility: WCAG 2.1 AA target (recommended)
* Role-specific views to reduce clutter (evaluators see assigned work first)

### Interoperability

* Import/export via:

  * Excel (XLSX)
  * CSV
  * PDF reports
* API (optional but valuable):

  * Read/write test definitions and results
  * Evidence metadata
  * Webhooks on approval/publish

### Observability

* Centralized logging
* Usage metrics (active users, completion rates)
* Error monitoring and alerting

---

## 6) MVP vs “Next” (so you can phase delivery)

### MVP (must have)

* Suite import from Excel + category/test browsing
* Project + vendor setup
* Test execution + scoring + evidence upload
* Weighted scoring + category rollups
* Vendor comparison + PDF export
* RBAC + audit logs

### Next (nice to have)

* Vendor portal for evidence submission
* Advanced workflow automation (approvals per category)
* APIs + webhooks
* Advanced analytics (trend across projects, benchmarking)
* Integrations (SharePoint/Drive evidence links, Jira tasks)

---

## 7) Acceptance criteria (examples)

* Import your suite Excel with >95% of rows mapped without manual editing
* Can evaluate **3 vendors** across **150–250 tests** with:

  * filtering by category/tier/priority
  * bulk assignment
  * evidence completeness tracking
* Generate a PDF scorecard showing:

  * overall weighted score
  * category breakdown
  * list of gating failures
* Audit log shows all score changes with timestamp + user

---

If you want, I can also produce this as:

* a **PRD format** (problem, users, journeys, requirements, success metrics), or
* a **user story + acceptance criteria backlog** ready for Jira (epics per module).
