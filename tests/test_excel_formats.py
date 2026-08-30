from task_management.common.excel_formats import number_format_for_header


def test_date_only_headers_use_readable_date_format() -> None:
    for header in ("登録日", "予定開始日", "期限", "完了日", "週"):
        assert number_format_for_header(header) == "yyyy/m/d"


def test_datetime_header_keeps_time() -> None:
    assert number_format_for_header("更新日時") == "yyyy/m/d h:mm"


def test_non_date_header_is_ignored() -> None:
    assert number_format_for_header("タスク名") is None
