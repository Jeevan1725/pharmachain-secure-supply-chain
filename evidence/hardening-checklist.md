# PharmaChain Hardening Checklist (Phase 15)
============================================

## Access Control
[x] SSH key-only authentication (no password)
[x] Root login disabled on all hosts
[x] MFA required for admin console access
[x] Least-privilege IAM roles; no wildcard permissions
[x] Segregation of duties across orgs
[x] Session timeout: 15 minutes idle

## Network / Ports / Services
[x] Only 443 exposed externally (API Gateway)
[x] Internal services on private subnet only
[x] Database not reachable from internet
[x] WAF in front of API Gateway
[x] NetworkPolicy allow-list between K8s pods
[x] Unused services disabled (telnet, ftp, etc.)

## Secrets
[x] All secrets in Vault (never in source, env, or config)
[x] Short-lived tokens only
[x] Automatic rotation every 90 days
[x] Audit log for every secret read

## Updates / Patches
[x] Weekly OS patch cycle
[x] Monthly dependency upgrade
[x] CVE monitoring via pip-audit + Dependabot
[x] Critical CVE SLA: fix within 72h

## Permissions / Files
[x] Non-root execution in containers
[x] Read-only root filesystem
[x] Resource limits on every pod
[x] Files owned by appuser:appuser

## Logging / Monitoring
[x] Security events forwarded to SIEM
[x] Tamper-evident log hashing
[x] 7-year retention for audit logs
[x] Alerts on authentication failures, ownership transfers, key rotations

## Physical Controls
[x] Data center biometric + badge access
[x] CCTV monitoring
[x] Locked server racks
[x] Redundant power (UPS + generator)
[x] Fire suppression

## Operational Controls
[x] Segregation of duties between Dev / Ops / Security
[x] Background checks for privileged staff
[x] Incident response plan documented
[x] Quarterly disaster recovery drills
[x] Change management approval required for production