from datetime import date

from task_management.common.operations import _date_header_groups


def test_date_header_groups_merges_consecutive_month_columns() -> None:
    columns = {
        11: date(2026, 7, 20),
        12: date(2026, 7, 27),
        13: date(2026, 8, 3),
        14: date(2026, 8, 10),
    }
    assert _date_header_groups(columns, "month") == [(11, 12, 7), (13, 14, 8)]


def test_date_header_groups_splits_year_boundary() -> None:
    columns = {11: date(2026, 12, 28), 12: date(2027, 1, 4)}
    assert _date_header_groups(columns, "year") == [(11, 11, 2026), (12, 12, 2027)]
