from __future__ import annotations

import logging

from scripts.run_full_pipeline import _PypdfWarningFilter


def _record(name: str, message: str, *args: object) -> logging.LogRecord:
    return logging.LogRecord(name, logging.WARNING, __file__, 1, message, args, None)


def test_repeated_wrong_pointer_warning_is_reported_once() -> None:
    warning_filter = _PypdfWarningFilter()
    first = _record(
        "pypdf._reader",
        "Ignoring wrong pointing object %d %d (offset %d)",
        2,
        65536,
        0,
    )
    repeated = _record(
        "pypdf._reader",
        "Ignoring wrong pointing object %d %d (offset %d)",
        15,
        65536,
        0,
    )

    assert warning_filter.filter(first) is True
    assert "이후 동일 경고는 생략합니다" in first.getMessage()
    assert warning_filter.filter(repeated) is False


def test_other_warnings_are_not_filtered() -> None:
    warning_filter = _PypdfWarningFilter()

    assert warning_filter.filter(_record("pypdf._reader", "another PDF warning")) is True
    assert (
        warning_filter.filter(
            _record("payroll_amendment", "Ignoring wrong pointing object 2 65536 (offset 0)")
        )
        is True
    )
