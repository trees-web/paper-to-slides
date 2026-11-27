#!/usr/bin/env python3
"""Assemble /app/output/result.json from a case's output contract.

Reads /input/case.json, extracts the output_contract field names (or any
`output_contract` / `outputContract` / `expected_outputs` / `expected_artifacts`
mapping), fills them from an explicit values file plus auto-detected artifacts in
the output directory, and appends honest status/sources/uncertainty/change records.

Usage:
  python write_result.py --case /input/case.json --out-dir /app/output \
      [--values values.json] [--status ok|partial|failed] \
      [--sources sources.json] [--uncertainty uncertainty.json] \
      [--change-record change.json]
"""
import argparse
import datetime
import json
import os
import re
import sys


def deep_find(obj, keys, depth=0):
    """Return first matching value for any key in `keys`, searching recursively."""
    if depth > 6:
        return None
    if isinstance(obj, dict):
        for k, v in obj.items():
            if k in keys:
                return v
            r = deep_find(v, keys, depth + 1)
            if r is not None:
                return r
    elif isinstance(obj, list):
        for v in obj:
            r = deep_find(v, keys, depth + 1)
            if r is not None:
                return r
    return None


def contract_keys(case):
    contract = deep_find(case, {"output_contract", "outputContract",
                                "expected_outputs", "expectedOutputs",
                                "expected_artifacts", "expectedArtifacts"})
    if isinstance(contract, dict):
        return list(contract.keys())
    return []


def list_files(out_dir):
    rel = []
    for root, _dirs, files in os.walk(out_dir):
        if ".lo-profile" in root.replace("\\", "/").split("/"):
            continue
        for f in files:
            if f.startswith(".~lock"):
                continue
            p = os.path.join(root, f)
            rel.append(os.path.relpath(p, out_dir).replace("\\", "/"))
    return sorted(rel)


def auto_fill(key, files):
    kl = key.lower()
    for f in files:
        fl = f.lower()
        if kl.endswith((".pptx", ".ppt", ".odp", ".pdf", ".png", ".jpg", ".json", ".md", ".txt")):
            base = os.path.splitext(fl)[0]
        else:
            base = fl
        if any(tok in kl for tok in ("deck", "slide", "ppt", "presentation", "演示", "幻灯片")) and fl.endswith((".pptx", ".ppt", ".odp")):
            return f
        if ("evidence" in kl or "证据" in kl or "mapping" in kl or "映射" in kl) and "evidence" in fl and fl.endswith(".json"):
            return f
        if ("outline" in kl or "plan" in kl or "大纲" in kl or "主线" in kl) and "outline" in fl and fl.endswith(".json"):
            return f
        if ("note" in kl or "script" in kl or "讲稿" in kl or "备注" in kl) and any(t in fl for t in ("notes", "script", "讲稿")) and fl.endswith((".json", ".md", ".txt")):
            return f
        if ("pdf" in kl or "render" in kl or "渲染" in kl) and fl.endswith(".pdf"):
            return f
        if ("preview" in kl or "png" in kl or "预览" in kl) and fl.endswith(".png"):
            return f
        if kl in ("case_id", "caseid"):
            return None
    return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--case", required=True)
    ap.add_argument("--out-dir", required=True)
    ap.add_argument("--values", default=None)
    ap.add_argument("--status", default="partial")
    ap.add_argument("--sources", default=None)
    ap.add_argument("--uncertainty", default=None)
    ap.add_argument("--change-record", default=None)
    args = ap.parse_args()

    case = json.load(open(args.case, encoding="utf-8"))
    case_id = case.get("case_id") or case.get("caseId") or case.get("id")
    keys = contract_keys(case)
    values = json.load(open(args.values, encoding="utf-8")) if args.values and os.path.exists(args.values) else {}
    files = list_files(args.out_dir)

    result = {}
    for k in keys:
        if k in values:
            result[k] = values[k]
        else:
            auto = auto_fill(k, files)
            result[k] = auto if auto is not None else values.get(k)
    missing = [k for k in keys if result.get(k) is None]

    result["case_id"] = case_id
    result["status"] = args.status
    result["sources"] = json.load(open(args.sources, encoding="utf-8")) if args.sources and os.path.exists(args.sources) else {}
    result["uncertainty"] = json.load(open(args.uncertainty, encoding="utf-8")) if args.uncertainty and os.path.exists(args.uncertainty) else {}
    result["change_record"] = json.load(open(args.change_record, encoding="utf-8")) if args.change_record and os.path.exists(args.change_record) else []
    result["artifacts"] = files
    result["missing_contract_fields"] = missing
    result["produced_at"] = datetime.datetime.now(datetime.timezone.utc).isoformat()

    out_path = os.path.join(args.out_dir, "result.json")
    json.dump(result, open(out_path, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
    print(json.dumps({"case_id": case_id, "contract_fields": keys, "filled": len(keys) - len(missing),
                      "missing": missing, "result": out_path}, ensure_ascii=False))


if __name__ == "__main__":
    main()
