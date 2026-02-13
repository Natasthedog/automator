from deck.engine.waterfall.targets import _normalize_column_name


def test_normalize_column_name_accepts_list_values() -> None:
    assert _normalize_column_name(["Target", "Brand"]) == "targetbrand"


def test_normalize_column_name_handles_none() -> None:
    assert _normalize_column_name(None) == ""
