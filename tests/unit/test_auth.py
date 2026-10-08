"""Unit tests for AuthService (Phase 14 test evidence)."""
import pytest
from unittest.mock import MagicMock


class FakeUser:
    def __init__(self, uid, email, pwd_hash, mfa_secret="SECRET"):
        self.user_id = uid
        self.email = email
        self.password_hash = pwd_hash
        self.mfa_secret = mfa_secret


class FakeUserRepo:
    def __init__(self, users):
        self._users = {u.email: u for u in users}

    def find_by_email(self, email):
        return self._users.get(email)


class FakeMFA:
    def __init__(self, ok=True):
        self.ok = ok

    def verify(self, secret, code):
        return self.ok


class FakeAudit:
    def __init__(self):
        self.events = []

    def log(self, kind, data):
        self.events.append((kind, data))


def make_auth(users=None, mfa_ok=True):
    # Local import so tests run even without top-level installation
    import sys
    from pathlib import Path
    sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "src"))
    from pharmachain.security.auth import AuthService

    repo = FakeUserRepo(users or [])
    return AuthService(repo, FakeMFA(mfa_ok), FakeAudit()), repo


def test_hash_password_rejects_short():
    auth, _ = make_auth()
    with pytest.raises(ValueError):
        auth.hash_password("short")


def test_hash_password_ok():
    auth, _ = make_auth()
    h = auth.hash_password("a-very-strong-password-123")
    assert h.startswith("$2")


def test_login_rejects_bad_email_format():
    auth, _ = make_auth()
    res = auth.login("not-an-email", "password123456", "123456")
    assert res.success is False
    assert res.reason == "Invalid credentials"


def test_login_rejects_bad_mfa_format():
    auth, _ = make_auth()
    res = auth.login("a@b.com", "password123456", "abc")
    assert res.success is False


def test_login_unknown_user_does_not_reveal():
    auth, _ = make_auth()
    res = auth.login("ghost@example.com", "password123456", "123456")
    assert res.success is False
    assert res.reason == "Invalid credentials"


def test_login_success():
    # Build a valid user
    import sys
    from pathlib import Path
    sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "src"))
    from pharmachain.security.auth import AuthService

    auth0 = AuthService(FakeUserRepo([]), FakeMFA(), FakeAudit())
    pwd = "a-very-strong-password-123"
    hashed = auth0.hash_password(pwd)

    user = FakeUser("u1", "a@b.com", hashed)
    auth, _ = make_auth(users=[user])
    res = auth.login("a@b.com", pwd, "123456")
    assert res.success is True
    assert res.user_id == "u1"


def test_login_locks_after_max_failures():
    auth, _ = make_auth()
    for _ in range(auth.MAX_FAILURES):
        auth.login("lock@example.com", "badpass123456", "123456")
    res = auth.login("lock@example.com", "badpass123456", "123456")
    assert res.reason == "Account locked"
