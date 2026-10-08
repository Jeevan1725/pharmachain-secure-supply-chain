"""
PharmaChain - Authentication Module (Phase 12 secure coding).
"""
from dataclasses import dataclass
from datetime import datetime, timedelta
import hashlib
import hmac

import bcrypt


@dataclass
class AuthResult:
    success: bool
    user_id: str = None
    reason: str = ""


class AuthService:
    MAX_FAILURES = 5
    LOCKOUT_WINDOW = timedelta(minutes=15)

    def __init__(self, user_repo, mfa_service, audit_logger):
        self.user_repo = user_repo
        self.mfa_service = mfa_service
        self.audit = audit_logger
        self._failure_cache = {}

    @staticmethod
    def _validate_email(email):
        if not email or "@" not in email or len(email) > 254:
            raise ValueError("Invalid email")

    @staticmethod
    def _validate_mfa_code(code):
        if not code or not code.isdigit() or len(code) != 6:
            raise ValueError("Invalid MFA code")

    def _is_locked(self, email):
        now = datetime.utcnow()
        recent = [t for t in self._failure_cache.get(email, [])
                  if now - t < self.LOCKOUT_WINDOW]
        self._failure_cache[email] = recent
        return len(recent) >= self.MAX_FAILURES

    def _record_failure(self, email):
        self._failure_cache.setdefault(email, []).append(datetime.utcnow())

    def login(self, email, password, mfa_code):
        try:
            self._validate_email(email)
            self._validate_mfa_code(mfa_code)
        except ValueError:
            self.audit.log("AUTH_FAIL", {"email": email, "reason": "validation"})
            return AuthResult(False, reason="Invalid credentials")

        if self._is_locked(email):
            self.audit.log("AUTH_LOCKED", {"email": email})
            return AuthResult(False, reason="Account locked")

        user = self.user_repo.find_by_email(email)
        if user is None:
            self._record_failure(email)
            self.audit.log("AUTH_FAIL", {"email": email, "reason": "user_not_found"})
            return AuthResult(False, reason="Invalid credentials")

        if not self._verify_password(password, user.password_hash):
            self._record_failure(email)
            self.audit.log("AUTH_FAIL", {"email": email, "reason": "bad_password"})
            return AuthResult(False, reason="Invalid credentials")

        if not self.mfa_service.verify(user.mfa_secret, mfa_code):
            self._record_failure(email)
            self.audit.log("AUTH_FAIL", {"email": email, "reason": "bad_mfa"})
            return AuthResult(False, reason="Invalid credentials")

        self.audit.log("AUTH_SUCCESS", {"email": email, "user_id": user.user_id})
        return AuthResult(True, user_id=user.user_id, reason="Authenticated")

    @staticmethod
    def hash_password(password):
        """Hash a password with bcrypt. Enforces minimum length 12."""
        if not password or len(password) < 12:
            raise ValueError("Password must be at least 12 characters")
        # bcrypt only uses the first 72 bytes; truncate explicitly to avoid errors
        pwd_bytes = password.encode("utf-8")[:72]
        return bcrypt.hashpw(pwd_bytes, bcrypt.gensalt()).decode("utf-8")

    @staticmethod
    def _verify_password(password, password_hash):
        if not password or not password_hash:
            return False
        pwd_bytes = password.encode("utf-8")[:72]
        return bcrypt.checkpw(pwd_bytes, password_hash.encode("utf-8"))

    @staticmethod
    def _constant_time_compare(a, b):
        return hmac.compare_digest(
            hashlib.sha256(a.encode()).digest(),
            hashlib.sha256(b.encode()).digest(),
        )