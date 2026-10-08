"""Unit tests for OwnershipService (Phase 14 test evidence)."""
import pytest
from dataclasses import dataclass


@dataclass
class FakeProduct:
    product_id: str
    owner: str


class FakeRepo:
    def __init__(self, products=None):
        self._p = {p.product_id: p for p in (products or [])}

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


class FakeLedger:
    def __init__(self):
        self.entries = []

    def append(self, payload):
        self.entries.append(payload)
        return "hash_" + str(len(self.entries))


class FakeAuthz:
    def __init__(self, allow=True):
        self.allow = allow

    def can(self, actor, action, product_id):
        return self.allow


class FakeSigner:
    def __init__(self, ok=True):
        self.ok = ok

    def verify(self, actor, product_id, new_owner, signature):
        return self.ok


class FakeAudit:
    def __init__(self):
        self.events = []

    def log(self, kind, data):
        self.events.append((kind, data))


class FakeActor:
    def __init__(self, uid="u1", org="orgA"):
        self.id = uid
        self.org_id = org


def make_service(products=None, allow=True, sig_ok=True):
    import sys
    from pathlib import Path
    sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "src"))
    from pharmachain.services.ownership_service import OwnershipService

    repo = FakeRepo(products)
    return OwnershipService(
        repo=repo,
        ledger=FakeLedger(),
        authz=FakeAuthz(allow),
        signer=FakeSigner(sig_ok),
        audit=FakeAudit(),
    )


def test_transfer_rejects_unauthorized():
    from pharmachain.services.ownership_service import UnauthorizedError
    svc = make_service(products=[FakeProduct("P1", "orgA")], allow=False)
    with pytest.raises(UnauthorizedError):
        svc.transfer("P1", "orgB", FakeActor(), b"sig")


def test_transfer_rejects_bad_signature():
    from pharmachain.services.ownership_service import SignatureError
    svc = make_service(products=[FakeProduct("P1", "orgA")], sig_ok=False)
    with pytest.raises(SignatureError):
        svc.transfer("P1", "orgB", FakeActor(), b"sig")


def test_transfer_rejects_unknown_product():
    from pharmachain.services.ownership_service import ProductNotFound
    svc = make_service(products=[], allow=True, sig_ok=True)
    with pytest.raises(ProductNotFound):
        svc.transfer("MISSING", "orgB", FakeActor(), b"sig")


def test_transfer_success():
    product = FakeProduct("P1", "orgA")
    svc = make_service(products=[product])
    result = svc.transfer("P1", "orgB", FakeActor(), b"sig")
    assert result.success is True
    assert result.tx_hash.startswith("hash_")
    assert product.owner == "orgB"
