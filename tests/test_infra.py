from payroll_amendment.infra import format_resident_id


def test_format_resident_id_adds_hyphen_and_preserves_mask() -> None:
    assert format_resident_id("9001011234567") == "900101-1234567"
    assert format_resident_id("900101-1******") == "900101-1******"


def test_format_resident_id_restores_numeric_leading_zero() -> None:
    assert format_resident_id(101011234567) == "010101-1234567"
