import csv
from unittest.mock import patch

import openpyxl
import pytest

import delete_transactions as dt


def make_review_workbook(path, rows):
    """rows: list of dicts with url, dated_on, amount, description,
    match_status, approve_delete."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Account A (Polluted)"
    headers = ["url", "dated_on", "amount", "description", "match_status", "approve_delete"]
    ws.append(headers)
    for r in rows:
        ws.append([r.get(h, "") for h in headers])
    wb.save(path)


class TestLoadApprovedRows:
    def test_filters_to_matched_and_approved_only(self, tmp_path):
        path = tmp_path / "review.xlsx"
        make_review_workbook(path, [
            dict(url="u1", dated_on="2026-01-01", amount=-10, description="a", match_status="matched", approve_delete="Y"),
            dict(url="u2", dated_on="2026-01-02", amount=-20, description="b", match_status="matched", approve_delete="N"),
            dict(url="u3", dated_on="2026-01-03", amount=-30, description="c", match_status="unmatched", approve_delete="Y"),
            dict(url="u4", dated_on="2026-01-04", amount=-40, description="d", match_status="AMBIGUOUS - review", approve_delete="Y"),
            dict(url="u5", dated_on="2026-01-05", amount=-50, description="e", match_status="matched", approve_delete="y"),
            dict(url="u6", dated_on="2026-01-06", amount=-60, description="f", match_status="matched", approve_delete=""),
        ])
        rows = dt.load_approved_rows(str(path))
        urls = {r["url"] for r in rows}
        assert urls == {"u1", "u5"}, "only exactly-matched rows explicitly approved with Y/y should pass"

    def test_empty_sheet_returns_empty_list(self, tmp_path):
        path = tmp_path / "review.xlsx"
        make_review_workbook(path, [])
        assert dt.load_approved_rows(str(path)) == []


def _row(url="u1", dated_on="2026-01-01", amount=-10.0, description="test"):
    return {"url": url, "dated_on": dated_on, "amount": amount, "description": description}


class TestProcess:
    def test_unexplained_transaction_dry_run_logs_delete_only_no_calls_made(self, tmp_path):
        log_path = tmp_path / "log.csv"
        with patch("delete_transactions.fa.get_transaction", return_value={"bank_transaction_explanations": []}), \
             patch("delete_transactions.fa.delete_transaction") as mock_del_txn, \
             patch("delete_transactions.fa.delete_explanation") as mock_del_exp:
            dt.process([_row()], live=False, log_path=str(log_path))

        mock_del_txn.assert_not_called()
        mock_del_exp.assert_not_called()
        rows = list(csv.DictReader(open(log_path)))
        assert len(rows) == 1
        assert rows[0]["action"] == "delete_transaction"
        assert rows[0]["result"] == "DRY_RUN"

    def test_explained_deletable_live_removes_explanation_then_deletes_transaction(self, tmp_path):
        log_path = tmp_path / "log.csv"
        txn = {"bank_transaction_explanations": [{"url": "exp1"}]}
        with patch("delete_transactions.fa.get_transaction", return_value=txn), \
             patch("delete_transactions.fa.get_explanation", return_value={"type": "Payment", "is_deletable": True}), \
             patch("delete_transactions.fa.delete_explanation") as mock_del_exp, \
             patch("delete_transactions.fa.delete_transaction") as mock_del_txn:
            dt.process([_row()], live=True, log_path=str(log_path))

        mock_del_exp.assert_called_once_with("exp1")
        mock_del_txn.assert_called_once()
        rows = list(csv.DictReader(open(log_path)))
        assert len(rows) == 1
        assert rows[0]["action"] == "remove_explanation + delete_transaction"
        assert rows[0]["result"] == "OK"
        assert rows[0]["detail"] == "Payment"

    def test_not_deletable_explanation_skips_transaction_entirely(self, tmp_path):
        log_path = tmp_path / "log.csv"
        txn = {"bank_transaction_explanations": [{"url": "exp1"}]}
        with patch("delete_transactions.fa.get_transaction", return_value=txn), \
             patch("delete_transactions.fa.get_explanation",
                   return_value={"type": "Payment", "is_deletable": False, "locked_reason": "VAT return filed"}), \
             patch("delete_transactions.fa.delete_explanation") as mock_del_exp, \
             patch("delete_transactions.fa.delete_transaction") as mock_del_txn:
            dt.process([_row()], live=True, log_path=str(log_path))

        mock_del_exp.assert_not_called()
        mock_del_txn.assert_not_called()
        rows = list(csv.DictReader(open(log_path)))
        assert len(rows) == 1
        assert rows[0]["result"] == "SKIPPED"
        assert "VAT return filed" in rows[0]["detail"]

    def test_invoice_receipt_is_skipped_entirely_even_if_deletable(self, tmp_path):
        log_path = tmp_path / "log.csv"
        txn = {"bank_transaction_explanations": [{"url": "exp1"}]}
        with patch("delete_transactions.fa.get_transaction", return_value=txn), \
             patch("delete_transactions.fa.get_explanation",
                   return_value={"type": "Invoice Receipt", "is_deletable": True}), \
             patch("delete_transactions.fa.delete_explanation") as mock_del_exp, \
             patch("delete_transactions.fa.delete_transaction") as mock_del_txn:
            dt.process([_row()], live=True, log_path=str(log_path))

        mock_del_exp.assert_not_called()
        mock_del_txn.assert_not_called()
        rows = list(csv.DictReader(open(log_path)))
        assert rows[0]["result"] == "SKIPPED"
        assert "Invoice Receipt" in rows[0]["detail"]

    def test_fetch_failure_is_logged_and_does_not_stop_remaining_rows(self, tmp_path):
        log_path = tmp_path / "log.csv"
        with patch("delete_transactions.fa.get_transaction", side_effect=RuntimeError("network down")), \
             patch("delete_transactions.fa.delete_transaction") as mock_del_txn:
            dt.process([_row(url="u1"), _row(url="u2")], live=True, log_path=str(log_path))

        mock_del_txn.assert_not_called()
        rows = list(csv.DictReader(open(log_path)))
        assert len(rows) == 2
        assert all(r["result"] == "ERROR" for r in rows)

    def test_authentication_failure_stops_before_retrying_next_row(self, tmp_path):
        log_path = tmp_path / "log.csv"
        with patch(
            "delete_transactions.fa.get_transaction",
            side_effect=dt.fa.AuthenticationError("token refresh failed"),
        ) as mock_get:
            dt.process([_row(url="u1"), _row(url="u2")], live=True, log_path=str(log_path))

        mock_get.assert_called_once_with("u1")
        rows = list(csv.DictReader(open(log_path)))
        assert len(rows) == 1
        assert rows[0]["result"] == "AUTH_ERROR"

    def test_dry_run_never_calls_live_functions_even_when_explained(self, tmp_path):
        log_path = tmp_path / "log.csv"
        txn = {"bank_transaction_explanations": [{"url": "exp1"}]}
        with patch("delete_transactions.fa.get_transaction", return_value=txn), \
             patch("delete_transactions.fa.get_explanation", return_value={"type": "Payment", "is_deletable": True}), \
             patch("delete_transactions.fa.delete_explanation") as mock_del_exp, \
             patch("delete_transactions.fa.delete_transaction") as mock_del_txn:
            dt.process([_row()], live=False, log_path=str(log_path))

        mock_del_exp.assert_not_called()
        mock_del_txn.assert_not_called()
        rows = list(csv.DictReader(open(log_path)))
        assert len(rows) == 1
        assert rows[0]["action"] == "remove_explanation + delete_transaction"
        assert rows[0]["result"] == "DRY_RUN"
        assert rows[0]["detail"] == "Payment"

    def test_explanation_deletion_error_stops_before_deleting_transaction(self, tmp_path):
        log_path = tmp_path / "log.csv"
        txn = {"bank_transaction_explanations": [{"url": "exp1"}]}
        with patch("delete_transactions.fa.get_transaction", return_value=txn), \
             patch("delete_transactions.fa.get_explanation", return_value={"type": "Payment", "is_deletable": True}), \
             patch("delete_transactions.fa.delete_explanation", side_effect=RuntimeError("API error")), \
             patch("delete_transactions.fa.delete_transaction") as mock_del_txn:
            dt.process([_row()], live=True, log_path=str(log_path))

        mock_del_txn.assert_not_called()
        rows = list(csv.DictReader(open(log_path)))
        assert rows[-1]["action"] == "remove_explanation"
        assert rows[-1]["result"] == "ERROR"


class TestMainConfirmationGate:
    """The live-run confirmation is the one thing standing between an
    approved list and irreversible deletion — worth locking in explicitly
    so a future refactor can't accidentally loosen it."""

    def test_anything_other_than_exact_DELETE_aborts(self, tmp_path, monkeypatch, capsys):
        path = tmp_path / "review.xlsx"
        make_review_workbook(path, [
            dict(url="u1", dated_on="2026-01-01", amount=-10, description="a", match_status="matched", approve_delete="Y"),
        ])
        monkeypatch.setattr("sys.argv", ["delete_transactions.py", str(path), "--live"])
        monkeypatch.setattr("builtins.input", lambda _: "yes please")
        with patch("delete_transactions.process") as mock_process:
            dt.main()
        mock_process.assert_not_called()
        assert "Aborted" in capsys.readouterr().out

    def test_exact_DELETE_proceeds(self, tmp_path, monkeypatch):
        path = tmp_path / "review.xlsx"
        make_review_workbook(path, [
            dict(url="u1", dated_on="2026-01-01", amount=-10, description="a", match_status="matched", approve_delete="Y"),
        ])
        monkeypatch.setattr("sys.argv", ["delete_transactions.py", str(path), "--live"])
        monkeypatch.setattr("builtins.input", lambda _: "DELETE")
        with patch("delete_transactions.process") as mock_process:
            dt.main()
        mock_process.assert_called_once()

    def test_dry_run_never_prompts(self, tmp_path, monkeypatch):
        path = tmp_path / "review.xlsx"
        make_review_workbook(path, [
            dict(url="u1", dated_on="2026-01-01", amount=-10, description="a", match_status="matched", approve_delete="Y"),
        ])
        monkeypatch.setattr("sys.argv", ["delete_transactions.py", str(path)])

        def _fail_if_called(_):
            raise AssertionError("dry run should not prompt for confirmation")

        monkeypatch.setattr("builtins.input", _fail_if_called)
        with patch("delete_transactions.process") as mock_process:
            dt.main()
        mock_process.assert_called_once()

    def test_no_approved_rows_skips_confirmation_and_processing(self, tmp_path, monkeypatch):
        path = tmp_path / "review.xlsx"
        make_review_workbook(path, [
            dict(url="u1", dated_on="2026-01-01", amount=-10, description="a", match_status="unmatched", approve_delete="Y"),
        ])
        monkeypatch.setattr("sys.argv", ["delete_transactions.py", str(path), "--live"])

        def _fail_if_called(_):
            raise AssertionError("should not prompt when there is nothing to do")

        monkeypatch.setattr("builtins.input", _fail_if_called)
        with patch("delete_transactions.process") as mock_process:
            dt.main()
        mock_process.assert_not_called()
