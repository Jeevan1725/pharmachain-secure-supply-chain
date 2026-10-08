# Phase 1 - Agile Process and Development Approach

## 1.1 Selected Approach: Scrum-XP

**Chosen:** Scrum with embedded XP (Extreme Programming) practices.

### Justification
PharmaChain involves five independent organizations operating under FDA DSCSA
regulatory constraints. Scrum iterative sprints allow phased cross-org
onboarding, while XP engineering practices (TDD, pair programming, CI, code
review) provide the rigor required for a security-critical, auditable system.

- Sprint length: 2 weeks
- Roles: Product Owner, Scrum Master, Dev Team (4), Security Champion
- Ceremonies: Planning, Daily Scrum, Review, Retrospective, Refinement
- XP practices: TDD, Pair Programming, CI, Small Releases, Simple Design

## 1.2 Agile Manifesto Mapping

| # | Principle | Application |
|---|-----------|-------------|
| 1 | Individuals & interactions | Cross-org daily scrums |
| 2 | Working software | Deployable module each sprint |
| 3 | Customer collaboration | Retailer + patient feedback in review |
| 4 | Responding to change | New FDA rule becomes new story |
| 5 | Continuous delivery | CI/CD ships signed images |

## 1.3 Refactoring Evidence

### Refactoring 1: Ownership Transfer (Monolithic to Layered)
- Before: 15-line function with SQLi, no auth, no audit
- After: OwnershipService class with RBAC, signature, ledger, audit
- Benefit: Eliminated SQLi, added non-repudiation, improved testability

### Refactoring 2: Serial Uniqueness Validation (DRY)
- Before: duplicated in 4 services
- After: single SerialValidator class
- Benefit: Single point of change, consistent errors, testable

## 1.4 Limitations & Mitigations

| Limitation | Mitigation |
|------------|------------|
| Under-specified security in sprints | Abuse stories + DoD with STRIDE review |
| Informal cross-org communication | Frozen OpenAPI contracts, CI schema validation |
| Compliance documentation pressure | Auto-generated Compliance Traceability Matrix |
| Velocity pressure skipping security | SAST + dependency scan mandatory for DoD |
| Governance ambiguity | Steering Committee as Product Owner delegate |

## 1.5 Traceability Thread (Carried Forward)

**SR-03** - Every ownership transfer must be digitally signed, hash-chained,
and audit-logged.
