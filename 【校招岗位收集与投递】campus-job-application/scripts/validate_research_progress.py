"""Validate private research progress without inferring external evidence truth."""
import json
import sys
from preference_contract import validate_snapshot

RESULTS = {"已有相关岗位记录", "无当年校招计划", "无用户要求岗位", "无法访问校招页面", "尚未推进"}
BUCKETS = {"first-pass-covered", "first-pass-visited-but-unresolved", "cycle/evidence-ambiguous-deferred", "first-pass-unvisited"}


def validate(data):
    errors = []
    records = data.get("records") if isinstance(data, dict) else None
    if not isinstance(records, list) or not records:
        return ["records must be a nonempty list"]
    errors.extend(validate_snapshot(data.get("preferences"), data.get("roles"), records, data.get("prior_roles"), batch_id=data.get("batch_id")))
    if not isinstance(data.get("batch_id"), str) or not data["batch_id"].strip():
        errors.append("batch_id required for batch default order")
    prefs = data.get("preferences")
    version = prefs.get("version") if isinstance(prefs, dict) else None
    seen = set()
    for index, row in enumerate(records):
        if not isinstance(row, dict):
            errors.append(f"record {index}: object required")
            continue
        if row.get("preference_version") != version or not version:
            errors.append(f"record {index}: stale preference_version")
        ident = row.get("id")
        prefix = f"record {index}"
        def error(message):
            errors.append(prefix + ": " + message)
        if not isinstance(ident, str) or not ident.strip() or ident in seen:
            error("unique nonempty id required")
        if isinstance(ident, str):
            seen.add(ident)
        for field in ["raw_status", "result_status", "coverage_bucket", "checked_at", "cycle_ref", "merge_status"]:
            if field not in row:
                error("missing " + field)
        if row.get("result_status") is not None and row.get("result_status") not in RESULTS:
            error("nonstandard result_status")
        for field in ["attempt_refs", "coverage_refs", "remaining_gaps"]:
            values = row.get(field)
            if not isinstance(values, list) or any(not isinstance(v, str) or not v.strip() for v in values):
                error(field + " must be a list of nonempty strings")
        bucket = row.get("coverage_bucket")
        if bucket not in BUCKETS:
            error("unknown coverage_bucket")
        attempts = row.get("attempt_refs", [])
        if bucket in {"first-pass-covered", "first-pass-visited-but-unresolved"}:
            if not attempts or not row.get("checked_at") or not row.get("cycle_ref"):
                error("visited requires dated current-cycle attempt evidence")
        if bucket == "first-pass-covered" and (not row.get("coverage_refs") or row.get("remaining_gaps")):
            error("covered requires verification references and no scoped gaps")
        if bucket == "first-pass-visited-but-unresolved" and not row.get("remaining_gaps"):
            error("unresolved requires explicit gaps")
        if bucket == "cycle/evidence-ambiguous-deferred" and (not row.get("remaining_gaps") or not attempts):
            error("ambiguous requires legacy evidence reference and uncertainty")
        if bucket == "first-pass-unvisited" and (attempts or row.get("coverage_refs")):
            error("unvisited contradicts documented attempts or coverage")
        if row.get("merge_status") not in {"not_merged", "partially_merged", "merged"}:
            error("unknown merge_status")
        if row.get("merge_status") == "merged" and not row.get("readback_ref"):
            error("merged requires saved tracker readback")
    return errors


if __name__ == "__main__":
    try:
        data = json.loads(open(sys.argv[1], encoding="utf-8").read())
        errors = validate(data)
    except (OSError, ValueError, IndexError) as exc:
        errors = [str(exc)]
    print(json.dumps({"passed": not errors, "errors": errors, "scope": "structure only; external evidence requires review"}, ensure_ascii=False))
    sys.exit(bool(errors))
