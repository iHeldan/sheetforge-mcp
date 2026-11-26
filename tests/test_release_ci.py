"""Offline checks for the read-only, exact-tag publication gate."""

import copy
import importlib.util
import json
from pathlib import Path
import runpy
import subprocess
import sys
from types import SimpleNamespace

import pytest

ROOT = Path(__file__).resolve().parents[1]
SHA = "a" * 40
REPO = "owner/sheetforge-mcp"


@pytest.fixture
def gate(monkeypatch):
    spec = importlib.util.spec_from_file_location("release_verifier", ROOT / "scripts/verify_release_artifacts.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    tag = f"v{module._package_version(ROOT)}"
    for name, value in {
        "GITHUB_REF_TYPE": "tag", "GITHUB_REF": f"refs/tags/{tag}",
        "GITHUB_REPOSITORY": REPO, "GITHUB_SHA": SHA,
    }.items():
        monkeypatch.setenv(name, value)
    run = {
        "id": 123, "run_number": 10, "run_attempt": 1,
        "head_sha": SHA, "head_branch": tag, "event": "push",
        "path": ".github/workflows/ci.yml", "head_repository": {"full_name": REPO},
        "status": "completed", "conclusion": "success",
    }
    jobs = []
    for name in sorted(module.REQUIRED_CI_JOBS):
        steps = [{"name": step, "status": "completed", "conclusion": "success"}
                 for step in module.REQUIRED_CI_STEPS]
        steps.append({"name": "Test installed wheel", "status": "completed",
                      "conclusion": "success" if name == "Lint and test (Python 3.12)" else "skipped"})
        jobs.append({"name": name, "run_id": 123, "run_attempt": 1, "head_sha": SHA,
                     "status": "completed", "conclusion": "success", "steps": steps})
    state = SimpleNamespace(
        module=module, run=run, runs=[[run]], jobs=[jobs], fresh=None, calls=[], error=None, error_at=1,
    )

    def request(command, **kwargs):
        state.calls.append(command)
        assert command[:7] == ["gh", "api", "--hostname", "github.com", "--method", "GET", command[6]]
        assert kwargs == {"check": True, "capture_output": True, "text": True, "timeout": 30}
        if state.error and len(state.calls) == state.error_at:
            raise state.error
        if command[6].endswith("/workflows/ci.yml/runs"):
            assert all(value in command for value in (f"head_sha={SHA}", f"branch={tag}", "event=push"))
            value = [{"workflow_runs": page} for page in state.runs]
        elif command[6].endswith("/jobs"):
            assert command[6] == f"repos/{REPO}/actions/runs/123/jobs"
            assert "filter=all" in command
            value = [{"jobs": page} for page in state.jobs]
        else:
            assert command[6] == f"repos/{REPO}/actions/runs/123"
            value = state.fresh if state.fresh is not None else state.run
        if not isinstance(value, dict):
            assert "--paginate" in command and "--slurp" in command and "per_page=100" in command
        return SimpleNamespace(stdout=json.dumps(value))

    monkeypatch.setattr(module.subprocess, "run", request)
    return state


def test_exact_tag_ci_passes_without_starting_any_workflows(gate, capsys):
    gate.module.verify_release_ci(ROOT)
    assert len(gate.calls) == 3
    assert "Verified release CI run 123 attempt 1" in capsys.readouterr().out


def test_real_cli_dispatches_the_ci_guard(gate, monkeypatch, capsys):
    script = ROOT / "scripts/verify_release_artifacts.py"
    monkeypatch.setattr(sys, "argv", [str(script), "--check-release-ci"])
    runpy.run_path(str(script), run_name="__main__")
    assert "Verified release CI run 123" in capsys.readouterr().out
    assert len(gate.calls) == 3


def test_required_jobs_and_steps_track_the_ci_matrix(gate):
    import yaml

    job = yaml.safe_load((ROOT / ".github/workflows/ci.yml").read_text())["jobs"]["lint-and-test"]
    matrix = job["strategy"]["matrix"]
    assert gate.module.REQUIRED_CI_JOBS == {
        *[f"Lint and test (Python {version})" for version in matrix["python-version"]],
        *[entry["name"] for entry in matrix["include"]],
    }
    for name in gate.module.REQUIRED_CI_STEPS:
        matches = [step for step in job["steps"] if step.get("name") == name]
        assert len(matches) == 1 and "if" not in matches[0]
        assert not matches[0].get("continue-on-error", False)


@pytest.mark.parametrize("field,value", [
    ("head_sha", "b" * 40), ("head_branch", "main"), ("event", "pull_request"),
    ("path", ".github/workflows/release-build.yml"),
    ("head_repository", {"full_name": "fork/sheetforge-mcp"}),
])
def test_other_commit_branch_workflow_or_fork_cannot_authorize_release(gate, field, value):
    gate.run[field] = value
    with pytest.raises(SystemExit, match="no push CI run"):
        gate.module.verify_release_ci(ROOT)
    assert len(gate.calls) == 1


def test_missing_run_fails_closed(gate):
    gate.runs = [[]]
    with pytest.raises(SystemExit, match="no push CI run"):
        gate.module.verify_release_ci(ROOT)


@pytest.mark.parametrize("status,conclusion", [
    ("queued", None), ("in_progress", None), ("waiting", None), ("requested", None),
    ("completed", "failure"), ("completed", "cancelled"), ("completed", "skipped"),
    ("completed", "timed_out"), ("completed", "neutral"), ("completed", "action_required"),
    ("completed", None), ("in_progress", "success"),
])
def test_incomplete_or_unsuccessful_run_is_rejected(gate, status, conclusion):
    gate.run.update(status=status, conclusion=conclusion)
    with pytest.raises(SystemExit, match="latest tag CI run is not successful"):
        gate.module.verify_release_ci(ROOT)
    assert len(gate.calls) == 1


@pytest.mark.parametrize("reverse", [False, True])
def test_newest_failure_cannot_fall_back_to_older_success(gate, reverse):
    newer = {**gate.run, "id": 456, "run_number": 11, "status": "queued", "conclusion": None}
    gate.runs = [[newer, gate.run] if reverse else [gate.run, newer]]
    with pytest.raises(SystemExit, match="latest tag CI run is not successful"):
        gate.module.verify_release_ci(ROOT)


@pytest.mark.parametrize("field", ["id", "run_number", "run_attempt"])
def test_incomplete_run_identity_is_rejected(gate, field):
    gate.run.pop(field)
    with pytest.raises(SystemExit, match="invalid CI run identity"):
        gate.module.verify_release_ci(ROOT)


@pytest.mark.parametrize("index", range(4))
def test_each_required_matrix_job_must_exist(gate, index):
    gate.jobs[0].pop(index)
    with pytest.raises(SystemExit, match="Linux/Windows jobs are missing"):
        gate.module.verify_release_ci(ROOT)


@pytest.mark.parametrize("index", range(4))
@pytest.mark.parametrize("conclusion", ["failure", "skipped", "cancelled"])
def test_green_workflow_cannot_mask_unsuccessful_required_job(gate, index, conclusion):
    gate.jobs[0][index]["conclusion"] = conclusion
    with pytest.raises(SystemExit, match="required job"):
        gate.module.verify_release_ci(ROOT)


@pytest.mark.parametrize("step", [
    "Install dependencies", "Verify Python runtime", "Run Ruff", "Run pytest",
    "Build package", "Verify release artifacts", "Test installed wheel",
])
@pytest.mark.parametrize("failure", ["missing", "skipped", "failure", "duplicate"])
def test_required_steps_and_installed_wheel_must_really_pass(gate, step, failure):
    job = next(job for job in gate.jobs[0] if job["name"] == "Lint and test (Python 3.12)")
    target = next(item for item in job["steps"] if item["name"] == step)
    if failure == "missing":
        job["steps"].remove(target)
    elif failure == "duplicate":
        job["steps"].append(copy.deepcopy(target))
    else:
        target["conclusion"] = failure
    with pytest.raises(SystemExit, match="required step"):
        gate.module.verify_release_ci(ROOT)


@pytest.mark.parametrize("field,value", [("run_id", 999), ("head_sha", "b" * 40), ("run_attempt", 2)])
def test_jobs_cannot_be_borrowed_from_other_runs_commits_or_future_attempts(gate, field, value):
    gate.jobs[0][0][field] = value
    with pytest.raises(SystemExit, match="invalid CI job identity"):
        gate.module.verify_release_ci(ROOT)


def test_pagination_is_consumed_for_both_runs_and_jobs(gate):
    gate.runs = [[{**gate.run, "head_branch": "main", "run_number": 999}], [gate.run]]
    gate.jobs = [gate.jobs[0][:2], gate.jobs[0][2:]]
    gate.module.verify_release_ci(ROOT)
    assert len(gate.calls) == 3


@pytest.mark.parametrize("success", [True, False])
def test_partial_rerun_uses_latest_job_attempt_without_repeating_passed_jobs(gate, success):
    gate.run["run_attempt"] = 2
    old = next(job for job in gate.jobs[0] if "Windows" in job["name"])
    retry = {**copy.deepcopy(old), "run_attempt": 2, "conclusion": "success" if success else "failure"}
    old["conclusion"] = "failure" if success else "success"
    gate.jobs[0].append(retry)
    if success:
        gate.module.verify_release_ci(ROOT)
    else:
        with pytest.raises(SystemExit, match="required job"):
            gate.module.verify_release_ci(ROOT)


def test_duplicate_job_is_not_silently_overwritten(gate):
    gate.jobs[0].append(copy.deepcopy(gate.jobs[0][0]))
    with pytest.raises(SystemExit, match="duplicate required CI job"):
        gate.module.verify_release_ci(ROOT)


@pytest.mark.parametrize("field,value", [
    ("status", "in_progress"), ("conclusion", "failure"), ("run_attempt", 2),
    ("head_sha", "b" * 40), ("id", 456),
])
def test_rerun_or_drift_during_verification_stops_publication(gate, field, value):
    gate.fresh = {**gate.run, field: value}
    with pytest.raises(SystemExit, match="CI changed during verification"):
        gate.module.verify_release_ci(ROOT)


@pytest.mark.parametrize("field,value", [
    ("GITHUB_REPOSITORY", ""), ("GITHUB_REPOSITORY", "../other/repo"),
    ("GITHUB_SHA", ""), ("GITHUB_SHA", "main"), ("GITHUB_REF_TYPE", "branch"),
])
def test_invalid_publication_context_makes_no_api_requests(gate, monkeypatch, field, value):
    monkeypatch.setenv(field, value)
    with pytest.raises(SystemExit):
        gate.module.verify_release_ci(ROOT)
    assert gate.calls == []


@pytest.mark.parametrize("error", [
    FileNotFoundError("gh missing"), subprocess.TimeoutExpired("gh", 30),
    subprocess.CalledProcessError(1, "gh", stderr="403 secret-must-not-be-logged"),
    subprocess.CalledProcessError(1, "gh", stderr="429 secret-must-not-be-logged"),
])
@pytest.mark.parametrize("request_number", [1, 2, 3])
def test_lookup_failure_is_closed_and_does_not_echo_credentials(gate, error, capsys, request_number):
    gate.error, gate.error_at = error, request_number
    with pytest.raises(SystemExit, match="actions: read") as caught:
        gate.module.verify_release_ci(ROOT)
    captured = capsys.readouterr()
    assert "secret-must-not-be-logged" not in str(caught.value) + captured.out + captured.err


def test_malformed_api_output_is_rejected(gate, monkeypatch):
    monkeypatch.setattr(gate.module.subprocess, "run", lambda *args, **kwargs: SimpleNamespace(stdout="not json"))
    with pytest.raises(SystemExit, match="lookup failed"):
        gate.module.verify_release_ci(ROOT)
