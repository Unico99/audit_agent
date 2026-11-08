import pandas as pd

from agent import checks
from agent.tools import run_tool


def _df(rows):
    return pd.DataFrame(
        rows, columns=["account_code", "account_name", "debit", "credit", "prior_balance"]
    )


def test_balanced():
    df = _df([("1", "نقدية", 100, 0, 90), ("2", "رأس مال", 0, 100, -90)])
    assert checks.check_balance(df)["balanced"] is True


def test_unbalanced_difference():
    df = _df([("1", "نقدية", 105, 0, 90), ("2", "رأس مال", 0, 100, -90)])
    r = checks.check_balance(df)
    assert r["balanced"] is False
    assert r["difference"] == 5.0


def test_duplicates_and_both_sides():
    df = _df([("1", "أ", 10, 5, 0), ("1", "ب", 0, 5, 0)])
    kinds = {i["issue"] for i in checks.check_data_quality(df)}
    assert "duplicate_account_code" in kinds
    assert "debit_and_credit_both_populated" in kinds


def test_non_numeric_detected():
    df = _df([("1", "أ", "abc", 0, 0), ("2", "ب", 0, 10, 0)])
    kinds = {i["issue"] for i in checks.check_data_quality(df)}
    assert "non_numeric_debit" in kinds


def test_analytical_flags_big_change():
    df = _df([("1", "عملاء", 150, 0, 100), ("2", "نقدية", 101, 0, 100)])
    flagged = checks.analytical_review(df, 20)["flagged"]
    assert [f["account_code"] for f in flagged] == ["1"]


def test_unknown_tool_returns_error():
    out, is_error = run_tool("nope", {})
    assert is_error and "error" in out


def test_path_traversal_blocked(tmp_path, monkeypatch):
    monkeypatch.setenv("AUDIT_DATA_DIR", str(tmp_path))
    out, is_error = run_tool("inspect_workbook", {"path": "../../etc/passwd"})
    assert is_error


def test_write_workpaper_needs_approval(tmp_path, monkeypatch):
    monkeypatch.setenv("AUDIT_OUTPUT_DIR", str(tmp_path))
    args = {"filename": "wp", "findings": [{"area": "TB", "finding": "x", "severity": "low", "recommendation": "y"}]}
    out, is_error = run_tool("write_workpaper", args, approver=lambda n, a: False)
    assert is_error and not list(tmp_path.glob("*.xlsx"))
    out, is_error = run_tool("write_workpaper", args, approver=lambda n, a: True)
    assert not is_error and (tmp_path / "wp.xlsx").exists()
