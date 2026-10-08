# PharmaChain Secure Deployment Checklist (Phase 15)
====================================================

## Pre-Deployment
[x] Signed container image (Cosign)
[x] K8s manifests reviewed (2-person)
[x] Secrets provisioned via Vault
[x] NetworkPolicy applied
[x] TLS certificates valid (>30 days)
[x] Resource limits set on all pods
[x] Rolling update strategy configured (maxSurge=1, maxUnavailable=0)

## Deployment
[x] Apply namespace + ConfigMap + Secret
[x] Apply Deployment (2 replicas)
[x] Apply Service (ClusterIP)
[x] Apply NetworkPolicy
[x] Verify pods Running + Ready
[x] Verify liveness + readiness probes passing

## Post-Deployment
[x] Smoke test: health endpoint 200
[x] Smoke test: authentication flow
[x] Smoke test: ownership transfer
[x] Verify logs flowing to SIEM
[x] Verify Prometheus metrics visible
[x] Verify alerts firing (dry run)

## Rollback Plan
[x] Previous image tagged and available
[x] Database migration rollback script ready
[x] kubectl rollout undo tested
[x] Postmortem template prepared