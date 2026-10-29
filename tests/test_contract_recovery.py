"""Guidance must repair request syntax without changing business definitions."""
# ruff: noqa: F811
# Imported pytest fixtures are intentionally injected into test arguments.

import copy
import json

import pytest

from excel_mcp import contracts, recipes, server
from excel_mcp.exceptions import (
    PreconditionFailedError,
    ResponseTooLargeError,
    ValidationError,
)
from tests.test_contracts import book, contract, recipe  # noqa: F401


def test_inspection_separates_logical_fields_from_headers(book, contract):
    report = contracts.inspect_workbook_contract(str(book), contract)
    assert report["recipe_input_fields"] == ["id", "net", "currency", "status"]
    assert report["bindings"]["net"] == "Net"
    assert report["next_action"]["tool"] == "run_workbook_recipe"
    assert report["next_action"]["mode"] == "preview"
    assert "binding keys" in report["field_usage"]


def test_exact_field_recovery_can_be_applied_and_committed(book, contract, recipe):
    recipe["group_by"] = ["Currency"]
    recipe["metrics"][0]["field"] = "Net"
    original = copy.deepcopy(recipe)
    before = book.read_bytes()
    reply = json.loads(
        server.run_workbook_recipe(str(book), recipe, workbook_contract=contract)
    )
    assert not reply["ok"] and book.read_bytes() == before and recipe == original
    error = reply["error"]
    assert (
        error["type"] == "ValidationError"
        and error["code"] == "contract_recipe_fields_invalid"
    )
    assert error["suggested_next_tool"] == "run_workbook_recipe"
    details = error["details"]
    assert details["retry_mode"] == "preview"
    assert details["invalid_reference_count"] == 2
    assert details["examples"] == [
        {
            "path": "/recipe/group_by/0",
            "provided": "Currency",
            "replace_with": "currency",
        },
        {"path": "/recipe/metrics/0/field", "provided": "Net", "replace_with": "net"},
    ]
    corrected = copy.deepcopy(recipe)
    corrected["group_by"][0] = details["examples"][0]["replace_with"]
    corrected["metrics"][0]["field"] = details["examples"][1]["replace_with"]
    preview = recipes.run_workbook_recipe(
        str(book), corrected, workbook_contract=contract
    )
    assert preview["rows_sample"] == [["EUR", 40]]
    assert recipes.run_workbook_recipe(
        str(book),
        corrected,
        workbook_contract=contract,
        mode="commit",
        expected_workbook_sha256=preview["expected_workbook_sha256"],
        recipe_token=preview["recipe_token"],
    )["persisted"]


def test_recovery_does_not_rewrite_literals_output_aliases_or_contract(
    book, contract, recipe
):
    recipe["filters"] = [{"field": "Status", "op": "eq", "value": "Net"}]
    recipe["metrics"][0]["as"] = "Net"
    original, approved = copy.deepcopy(recipe), copy.deepcopy(contract)
    with pytest.raises(ValidationError) as caught:
        recipes.run_workbook_recipe(str(book), recipe, workbook_contract=contract)
    assert caught.value.details["examples"] == [
        {
            "path": "/recipe/filters/0/field",
            "provided": "Status",
            "replace_with": "status",
        }
    ]
    assert recipe == original and contract == approved
    assert "Net" not in [
        item.get("replace_with") for item in caught.value.details["examples"]
    ]


def test_unknown_names_have_no_fuzzy_replacement(book, contract, recipe):
    recipe["metrics"][0]["field"] = "NET"
    with pytest.raises(ValidationError) as caught:
        recipes.run_workbook_recipe(str(book), recipe, workbook_contract=contract)
    example = caught.value.details["examples"][0]
    assert example["provided"] == "NET" and "replace_with" not in example
    assert caught.value.details["allowed_input_fields"] == [
        "id",
        "net",
        "currency",
        "status",
    ]


def test_valid_logical_name_is_not_replaced_by_a_physical_header(
    book, contract, recipe
):
    contract["fields"][0]["name"] = "Net"
    contract["key_fields"] = ["Net"]
    recipe["group_by"] = ["Net"]
    preview = recipes.run_workbook_recipe(str(book), recipe, workbook_contract=contract)
    assert preview["rows_sample"] == [[1, 10], [3, 30]]
    assert preview["headers"] == ["Net", "Revenue"]


def test_source_rules_fail_before_recipe_field_recovery(book, contract, recipe):
    contract["fields"][1]["min"] = 15
    recipe["metrics"][0]["field"] = "Net"
    before = book.read_bytes()
    with pytest.raises(PreconditionFailedError) as caught:
        recipes.run_workbook_recipe(str(book), recipe, workbook_contract=contract)
    assert caught.value.code == "workbook_contract_failed"
    assert caught.value.details["next_action"]["action"] == "ask_data_owner"
    assert "replace_with" not in json.dumps(caught.value.details)
    assert book.read_bytes() == before


def test_guidance_is_not_part_of_recipe_or_binding_tokens(
    book, contract, recipe, monkeypatch
):
    with_guidance = recipes.run_workbook_recipe(
        str(book), recipe, workbook_contract=contract
    )
    inspect = contracts.inspect_table_contract

    def without_guidance(*args):
        report = inspect(*args)
        for key in ("recipe_input_fields", "field_usage", "next_action"):
            report.pop(key)
        return report

    monkeypatch.setattr(contracts, "inspect_table_contract", without_guidance)
    without = recipes.run_workbook_recipe(str(book), recipe, workbook_contract=contract)
    assert with_guidance["recipe_token"] == without["recipe_token"]
    assert (
        with_guidance["workbook_contract"]["binding_sha256"]
        == without["workbook_contract"]["binding_sha256"]
    )


def test_check_source_references_have_exact_paths(book, contract, recipe):
    recipe["checks"] = [
        {"type": "unique_key", "fields": ["Id"]},
        {"type": "sum_matches", "source_field": "Net", "output_field": "Revenue"},
    ]
    with pytest.raises(ValidationError) as caught:
        recipes.run_workbook_recipe(str(book), recipe, workbook_contract=contract)
    assert caught.value.details["examples"] == [
        {"path": "/recipe/checks/0/fields/0", "provided": "Id", "replace_with": "id"},
        {
            "path": "/recipe/checks/1/source_field",
            "provided": "Net",
            "replace_with": "net",
        },
    ]


def test_repeated_recovery_examples_are_bounded(book, contract, recipe):
    recipe["filters"] = [{"field": "Net", "op": "gte", "value": 0} for _ in range(20)]
    with pytest.raises(ValidationError) as caught:
        recipes.run_workbook_recipe(str(book), recipe, workbook_contract=contract)
    assert caught.value.details["invalid_reference_count"] == 20
    assert len(caught.value.details["examples"]) == 5
    assert caught.value.details["examples_truncated"]


def test_recovery_budget_fails_before_any_commit(book, contract, recipe, monkeypatch):
    preview = recipes.run_workbook_recipe(str(book), recipe, workbook_contract=contract)
    before = book.read_bytes()
    recipe["metrics"][0]["field"] = "\u00e4" * 128
    report = contracts.inspect_table_contract

    def checked_source(*args):
        result = report(*args)
        # Isolate the recovery payload's budget from the source report budget.
        monkeypatch.setattr(contracts, "MAX_CONTRACT_SUMMARY", 10)
        return result

    monkeypatch.setattr(contracts, "inspect_table_contract", checked_source)
    with pytest.raises(ResponseTooLargeError):
        recipes.run_workbook_recipe(
            str(book),
            recipe,
            workbook_contract=contract,
            mode="commit",
            recipe_token=preview["recipe_token"],
            expected_workbook_sha256=preview["expected_workbook_sha256"],
        )
    assert book.read_bytes() == before


def test_source_recovery_copies_approved_selector_only(book, contract, recipe):
    recipe["source"]["sheet_name"] = "Data"
    approved = copy.deepcopy(contract)
    with pytest.raises(ValidationError) as caught:
        recipes.run_workbook_recipe(str(book), recipe, workbook_contract=contract)
    assert caught.value.code == "contract_recipe_source_mismatch"
    assert caught.value.details == {
        "path": "/recipe/source",
        "replace_with": {"table_name": "Sales"},
        "retry_mode": "preview",
    }
    assert contract == approved


@pytest.mark.parametrize("bad", [0, 6, 10, True, "3"])
def test_sample_limit_recovery_is_explicit(book, contract, bad):
    with pytest.raises(ValidationError) as caught:
        contracts.inspect_workbook_contract(str(book), contract, sample_limit=bad)
    assert caught.value.code == "contract_sample_limit_invalid"
    assert caught.value.details["suggested_arguments"] == {"sample_limit": 3}
    assert caught.value.suggested_next_tool == "inspect_workbook_contract"


def test_unresolved_bindings_require_owner_not_a_guess(book, contract):
    contract["fields"][1]["headers"] = ["Missing"]
    report = contracts.inspect_workbook_contract(str(book), contract)
    assert not report["passed"] and report["next_action"]["action"] == "ask_data_owner"
    assert report["checked"] == []
    assert "replace_with" not in json.dumps(report)


def test_violations_do_not_propose_filtering_away_source_rows(book, contract):
    contract["fields"][1]["min"] = 15
    report = contracts.inspect_workbook_contract(str(book), contract)
    assert (
        not report["passed"]
        and report["next_action"]["reason"] == "source_rules_failed"
    )
    assert "Do not weaken" in report["next_action"]["instruction"]


def test_ordinary_validation_error_keeps_legacy_envelope():
    error = ValidationError("old message")
    payload = json.loads(server._error_response("test", error))
    assert payload["error"] == {"type": "ValidationError", "message": "old message"}
    assert ValidationError("one", "two").args == ("one", "two")
