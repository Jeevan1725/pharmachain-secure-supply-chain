"""
Integration test: verify ownership transfer writes to ledger AND audit log
atomically (Saga pattern).
"""
import sys
from pathlib import Path
from dataclasses import dataclass

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "src"))


@dataclass
class Product:
    product_id: str
    owner: str


class InMemoryRepo:
    def __init__(self):
        self._p = {}
    def seed(self, product):
        self._p[product.product_id] = product
    def get(self, pid):
        return self._p.get(pid)
    def save(self, product):
        self._p[product.product_id] = product
    def transaction(self):
        class _T:
            def __enter__(self_): pass
            def __exit__(self_, *a): return False
        return _T()
    def find_by_serial(self, serial):
        return None


class InMemoryLedger:
    def __init__(self):
        self.entries = []
        self.prev_hash = "genesis"
    def append(self, payload):
        entry = dict(payload)
        entry["prev_hash"] = self.prev_hash
        entry["hash"] = "h_" + str(len(self.entries) + 1)
        self.entries.append(entry)
        self.prev_hash = entry["hash"]
        return entry["hash"]


class InMemoryAudit:
    def __init__(self):
        self.events = []
    def log(self, kind, data):
        self.events.append((kind, data))


class AllowAllAuthz:
    def can(self, actor, action, product_id):
        return True


class AlwaysValidSigner:
    def verify(self, actor, product_id, new_owner, signature):
        return True


class FakeActor:
    def __init__(self, uid, org):
        self.id = uid
        self.org_id = org


def test_integration_transfer_writes_to_both_ledger_and_audit():
    from pharmachain.services.ownership_service import OwnershipService

    repo = InMemoryRepo()
    repo.seed(Product("P-1", "Pfizer"))
    ledger = InMemoryLedger()
    audit = InMemoryAudit()

    svc = OwnershipService(
        repo=repo, ledger=ledger,
        authz=AllowAllAuthz(), signer=AlwaysValidSigner(), audit=audit,
    )

    result = svc.transfer("P-1", "MedLogix", FakeActor("u1", "Pfizer"), b"sig-xyz")

    assert repo.get("P-1").owner == "MedLogix"
    assert len(ledger.entries) == 1
    entry = ledger.entries[0]
    assert entry["type"] == "OWNERSHIP_TRANSFER"
    assert entry["from"] == "Pfizer"
    assert entry["to"] == "MedLogix"
    assert entry["signature"] == "7369672d78797a"
    assert entry["prev_hash"] == "genesis"
    kinds = [k for k, _ in audit.events]
    assert "TRANSFER" in kinds
    assert result.success
    assert result.tx_hash == "h_1"


def test_integration_chain_grows_across_multiple_transfers():
    from pharmachain.services.ownership_service import OwnershipService

    repo = InMemoryRepo()
    repo.seed(Product("P-1", "Pfizer"))
    ledger = InMemoryLedger()
    audit = InMemoryAudit()

    svc = OwnershipService(
        repo=repo, ledger=ledger,
        authz=AllowAllAuthz(), signer=AlwaysValidSigner(), audit=audit,
    )

    svc.transfer("P-1", "MedLogix", FakeActor("u1", "Pfizer"), b"sig1")
    svc.transfer("P-1", "Apollo",   FakeActor("u2", "MedLogix"), b"sig2")

    assert len(ledger.entries) == 2
    assert ledger.entries[0]["hash"] == ledger.entries[1]["prev_hash"]
    assert repo.get("P-1").owner == "Apollo"