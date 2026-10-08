"""
Fuzz test: Hypothesis sends random serial strings to the validator.
"""
import sys
from pathlib import Path
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "src"))

from hypothesis import given, strategies as st, settings


class DummyRepo:
    def find_by_serial(self, serial):
        return None


def test_serial_validator_never_crashes_on_random_input():
    from pharmachain.utils.validators import SerialValidator, ValidationError

    validator = SerialValidator(DummyRepo())

    @settings(max_examples=500, deadline=None)
    @given(st.text(max_size=200))
    def fuzz(raw):
        try:
            validator.validate_format(raw)
        except ValidationError:
            pass

    fuzz()


def test_serial_validator_accepts_valid_serials():
    from pharmachain.utils.validators import SerialValidator
    SerialValidator(DummyRepo()).validate_format("SN-ABCD-1234-EF56")


def test_serial_validator_rejects_empty():
    from pharmachain.utils.validators import SerialValidator, ValidationError
    with pytest.raises(ValidationError):
        SerialValidator(DummyRepo()).validate_format("")


def test_serial_validator_rejects_sql_injection():
    from pharmachain.utils.validators import SerialValidator, ValidationError
    with pytest.raises(ValidationError):
        SerialValidator(DummyRepo()).validate_format("SN-ABCD'; DROP TABLE products;--")