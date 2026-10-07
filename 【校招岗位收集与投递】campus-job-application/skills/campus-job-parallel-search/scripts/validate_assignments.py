#!/usr/bin/env python3
"""Validate a private coordination snapshot. No locks, browser actions or truth checks."""
import argparse
import json
import re
from pathlib import Path, PurePosixPath

# Resolve inside the complete installed parent package, independent of cwd.
import importlib.util
_spec = importlib.util.spec_from_file_location("preference_contract", Path(__file__).resolve().parents[3] / "scripts/preference_contract.py")
_contract = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_contract)

WORKFLOW = 'search'
STAGES = {"discovered", "verified", "preparing", "prepared", "reviewed", "submit_clicked", "submitted", "uncertain"}
PREPARED = {"prepared", "reviewed", "submit_clicked", "submitted", "uncertain"}


def text(value):
    return isinstance(value, str) and bool(value.strip())


def relative(value):
    return text(value) and "\\" not in value and ":" not in value and not value.startswith("/") and all(p not in {"", ".", ".."} for p in value.split("/"))


def within(path, root):
    return path == root or root in PurePosixPath(path).parents


def search_only_errors(role):
    errors = []
    if role.get("stage") not in {None, "discovered", "verified"}:
        errors.append("search role stage must remain discovered/verified (read-only)")
    execution = role.get("execution")
    if not isinstance(execution, dict) or any(
        not isinstance(execution.get(action), dict)
        or execution[action].get("status") != "not_started"
        for action in ("fill", "save", "submit")
    ):
        errors.append("search role fill/save/submit must all remain not_started")
    return errors


def validate(data):
    errors = []
    if not isinstance(data, dict):
        return ["manifest must be an object"]
    if data.get("workflow") != WORKFLOW:
        errors.append("wrong workflow")
    for field in ("batch_id", "query_version", "mode_approval_ref"):
        if not text(data.get(field)):
            errors.append("missing " + field)
    if not text(data.get("mode")) or data.get("mode") not in {"single", "parallel"}:
        errors.append("execution mode must be explicitly selected")
    if not text(data.get("execution_channel")) or data.get("execution_channel") not in {"platform", "manual_handoff"}:
        errors.append("missing execution channel")
    participants = data.get("participants")
    if not isinstance(participants, list) or not participants or any(not text(p) for p in participants):
        errors.append("participants must be nonempty distinct IDs")
        participants = []
    elif len(set(participants)) != len(participants):
        errors.append("duplicate participant")
    if data.get("mode") == "single" and len(participants) != 1:
        errors.append("single mode needs one participant")
    if data.get("mode") == "parallel" and len(participants) < 2:
        errors.append("parallel mode needs multiple participants")
    if data.get("coordinator") not in participants:
        errors.append("coordinator must be a participant")
    if not isinstance(data.get("source_hash"), str) or not re.fullmatch(r"[a-fA-F0-9]{64}", data["source_hash"]):
        errors.append("invalid source hash")
    if type(data.get("assignment_version")) is not int or data["assignment_version"] < 1:
        errors.append("invalid assignment version")
    roots = data.get("output_roots")
    if not isinstance(roots, dict):
        roots = {}
    if set(roots) != set(participants) or any(not relative(v) for v in roots.values()):
        errors.append("output roots must match participants and be relative")
    valid_roots = [PurePosixPath(v) for v in roots.values() if relative(v)]
    for i, a in enumerate(valid_roots):
        for b in valid_roots[i + 1:]:
            if within(a, b) or within(b, a):
                errors.append("participant output roots overlap")
    assignments = data.get("assignments")
    if not isinstance(assignments, list):
        errors.append("assignments must be a list")
        assignments = []
    ids, companies, scopes, entities = {}, set(), set(), set()
    scope_field = "source_scope" if WORKFLOW == "search" else "operation_scope"
    for item in assignments:
        if not isinstance(item, dict):
            errors.append("assignment must be an object")
            continue
        key = item.get("assignment_id")
        owner = item.get("owner")
        for field, seen in (("assignment_id", ids), ("company_key", companies), (scope_field, scopes)):
            value = item.get(field)
            if not text(value):
                errors.append("missing assignment " + field)
            elif value in seen:
                errors.append("duplicate allocation " + field)
            elif isinstance(seen, dict):
                seen[value] = item
            else:
                seen.add(value)
        if WORKFLOW == "application":
            entity = item.get("recruiting_entity")
            if not text(entity) or entity in entities:
                errors.append("missing or duplicate recruiting entity")
            else:
                entities.add(entity)
        if not text(owner) or owner not in participants:
            errors.append("unknown assignment owner")
        output = item.get("output_dir")
        root = roots.get(owner) if text(owner) else None
        if not relative(output) or not relative(root) or not within(PurePosixPath(output), PurePosixPath(root)):
            errors.append("assignment output outside owner root")
        handoff = item.get("handoff")
        if handoff is not None:
            if not isinstance(handoff, dict):
                errors.append("handoff must be an object")
            elif (handoff.get("from_owner") not in participants or handoff.get("to_owner") != owner or handoff.get("from_owner") == owner or not text(handoff.get("stopped_ref")) or not text(handoff.get("accepted_ref"))):
                errors.append("handoff needs old-owner stop and new-owner acknowledgement")
    roles = data.get("roles", [])
    if not isinstance(roles, list):
        errors.append("roles must be a list")
        roles = []
    seen_keys, seen_urls = set(), set()
    for role in roles:
        if not isinstance(role, dict):
            errors.append("role must be an object")
            continue
        for field, seen in (("role_key", seen_keys), ("direct_url", seen_urls)):
            value = role.get(field)
            if not text(value) or value in seen:
                errors.append("missing or duplicate role " + field)
            else:
                seen.add(value)
        assignment = ids.get(role.get("assignment_id")) if text(role.get("assignment_id")) else None
        if assignment is None or role.get("owner") != assignment.get("owner"):
            errors.append("role has unknown assignment or wrong owner")
        if WORKFLOW == "search":
            errors.extend(search_only_errors(role))
            if not text(role.get("jd_full_text")) or not text(role.get("jd_capture_ref")):
                errors.append("search role needs full JD and capture reference")
        else:
            stage = role.get("stage")
            if not text(stage):
                errors.append("role stage must be a string")
                continue
            if stage not in STAGES:
                errors.append("invalid role stage")
            if stage in PREPARED and not text(role.get("material_id")):
                errors.append("prepared role needs material identity")
            if stage in {"reviewed", "submit_clicked", "submitted"} and not text(role.get("review_hash")):
                errors.append("reviewed role needs current review hash")
            if stage in {"submit_clicked", "submitted"} and (not text(role.get("approval_ref")) or not text(role.get("review_hash")) or role.get("approved_review_hash") != role.get("review_hash")):
                errors.append("submission lacks approval for current review")
            if stage == "submitted" and not text(role.get("receipt_ref")):
                errors.append("submitted role needs official receipt")
            if stage != "submitted" and text(role.get("tracker_status")) and role.get("tracker_status") in {"submitted", "已提交"}:
                errors.append("tracker claims submitted without verified stage")
    errors.extend(_contract.validate_snapshot(data.get("preferences"), roles, assignments, data.get("prior_roles")))
    prefs = data.get("preferences")
    preference_version = prefs.get("version") if isinstance(prefs, dict) else None
    for assignment in assignments:
        if isinstance(assignment, dict) and (assignment.get("preference_version") != preference_version or not preference_version):
            errors.append("stale assignment preference_version")
    deltas = data.get("deltas", [])
    if not isinstance(deltas, list):
        errors.append("deltas must be a list")
        deltas = []
    for delta in deltas:
        if not isinstance(delta, dict):
            errors.append("delta must be an object")
            continue
        for field in ("batch_id", "query_version", "assignment_version", "source_hash"):
            if delta.get(field) != data.get(field):
                errors.append("stale or foreign delta " + field)
        if delta.get("preference_version") != preference_version or not preference_version:
            errors.append("stale delta preference_version")
        assignment = ids.get(delta.get("assignment_id")) if text(delta.get("assignment_id")) else None
        if assignment is None or delta.get("owner") != assignment.get("owner"):
            errors.append("delta has unknown assignment or wrong owner")
        if "roles" in delta:
            if isinstance(delta["roles"], list):
                for role in delta["roles"]:
                    if isinstance(role, dict):
                        errors.extend(search_only_errors(role))
            errors.extend(_contract.validate_snapshot(data.get("preferences"), delta["roles"], assignments, []))
        if delta.get("merge_status") == "merged" and not text(delta.get("tracker_readback_ref")):
            errors.append("merged delta needs tracker readback")
    return errors


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("manifest", type=Path)
    args = parser.parse_args()
    try:
        errors = validate(json.loads(args.manifest.read_text(encoding="utf-8")))
    except (OSError, ValueError) as exc:
        print(json.dumps({"valid": False, "errors": [str(exc)]}, ensure_ascii=False))
        return 2
    print(json.dumps({"valid": not errors, "errors": errors, "workflow": WORKFLOW}, ensure_ascii=False, indent=2))
    return int(bool(errors))


if __name__ == "__main__":
    raise SystemExit(main())
