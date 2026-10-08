"""
PharmaChain - Shared Validators (Phase 1 Refactoring 2: DRY).
"""
import re

_SERIAL_RE = re.compile(r"^SN-[A-F0-9]{4}-[A-F0-9]{4}-[A-F0-9]{4}$")


class DuplicateSerialError(Exception):
    pass


class ValidationError(Exception):
    pass


class SerialValidator:
    def __init__(self, repo):
        self.repo = repo

    def validate_format(self, serial):
        if not isinstance(serial, str) or not _SERIAL_RE.match(serial):
            raise ValidationError("Invalid serial format: " + str(serial))

    def ensure_unique(self, serial):
        self.validate_format(serial)
        if self.repo.find_by_serial(serial):
            raise DuplicateSerialError(serial)