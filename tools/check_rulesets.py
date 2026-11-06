#!/usr/bin/env python3
"""Compare the repository's live rulesets with the versioned copies.

The rulesets that protect main, the release branches and the release tags are
configured in GitHub settings. A change there leaves no trace in review, so
each ruleset is versioned as normalized JSON under .github/rulesets/, and this
check reports any difference between the live rulesets and those files.

Read-only: it never changes a ruleset. To change one, edit it in the GitHub
settings, then run this tool with --export and commit the result in the same
pull request that explains why.

Normalization keeps what a ruleset enforces (name, target, enforcement,
conditions, rules and their parameters) and drops what GitHub generates or
hides: ids, links, timestamps, and bypass actors, which the API returns only
to repository administrators.

Exit status: 0 when the live rulesets match the files, 1 on drift, 2 on a
usage, network or repository error.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
import urllib.error
import urllib.request
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
RULESET_DIR = Path(".github") / "rulesets"
API = "https://api.github.com"
KEPT = ("name", "target", "enforcement", "conditions", "rules")


def normalize(ruleset: dict[str, Any]) -> dict[str, Any]:
    """Return the enforced content of one ruleset in a stable form."""
    missing = [key for key in KEPT if key not in ruleset]
    if missing:
        raise ValueError(f"ruleset {ruleset.get('name')!r} lacks {', '.join(missing)}")
    rules = sorted(
        ({"type": rule["type"], **({"parameters": rule["parameters"]}
                                   if rule.get("parameters") is not None else {})}
         for rule in ruleset["rules"]),
        key=lambda rule: str(rule["type"]),
    )
    return {
        "name": ruleset["name"],
        "target": ruleset["target"],
        "enforcement": ruleset["enforcement"],
        "conditions": ruleset["conditions"],
        "rules": rules,
    }


def file_name(name: str) -> str:
    slug = re.sub(r"[^a-z0-9]+", "-", name.lower()).strip("-")
    if not slug:
        raise ValueError(f"ruleset name {name!r} has no usable characters")
    return f"{slug}.json"


def render(ruleset: dict[str, Any]) -> str:
    return json.dumps(ruleset, indent=2, sort_keys=True, ensure_ascii=False) + "\n"


def load_versioned(root: Path) -> dict[str, dict[str, Any]]:
    directory = root / RULESET_DIR
    versioned: dict[str, dict[str, Any]] = {}
    for path in sorted(directory.glob("*.json")):
        data = json.loads(path.read_text(encoding="utf-8"))
        if not isinstance(data, dict):
            raise ValueError(f"{path.relative_to(root)} is not a JSON object")
        ruleset = normalize(data)
        if data != ruleset:
            raise ValueError(f"{path.relative_to(root)} is not normalized; "
                             "regenerate it with --export")
        if path.name != file_name(ruleset["name"]):
            raise ValueError(f"{path.relative_to(root)} must be named "
                             f"{file_name(ruleset['name'])}")
        versioned[ruleset["name"]] = ruleset
    return versioned


def fetch(url: str, token: str) -> Any:
    headers = {"Accept": "application/vnd.github+json", "X-GitHub-Api-Version": "2022-11-28"}
    if token:
        headers["Authorization"] = f"Bearer {token}"
    request = urllib.request.Request(url, headers=headers)
    with urllib.request.urlopen(request, timeout=30) as response:
        return json.load(response)


def load_live(repository: str, token: str) -> dict[str, dict[str, Any]]:
    """Fetch every ruleset defined on the repository itself."""
    summaries = fetch(f"{API}/repos/{repository}/rulesets?includes_parents=false&per_page=100",
                      token)
    if not isinstance(summaries, list):
        raise ValueError("the rulesets endpoint did not return a list")
    if len(summaries) >= 100:
        raise ValueError("100 or more rulesets; refusing to compare a possibly partial list")
    live: dict[str, dict[str, Any]] = {}
    for summary in summaries:
        detail = fetch(f"{API}/repos/{repository}/rulesets/{int(summary['id'])}", token)
        ruleset = normalize(detail)
        live[ruleset["name"]] = ruleset
    return live


def compare(versioned: dict[str, dict[str, Any]],
            live: dict[str, dict[str, Any]]) -> list[str]:
    findings = []
    for name in sorted(set(versioned) | set(live)):
        if name not in live:
            findings.append(f"{name!r} is versioned but not defined on GitHub")
        elif name not in versioned:
            findings.append(f"{name!r} is defined on GitHub but not versioned in "
                            f"{RULESET_DIR.as_posix()}/{file_name(name)}")
        elif versioned[name] != live[name]:
            changed = [key for key in KEPT if versioned[name][key] != live[name][key]]
            findings.append(f"{name!r} differs from its versioned copy in: {', '.join(changed)}")
    return findings


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0] if __doc__ else None)
    parser.add_argument("--root", type=Path, default=ROOT)
    source = parser.add_mutually_exclusive_group(required=True)
    source.add_argument("--repository", help="owner/name; reads GITHUB_TOKEN if set")
    source.add_argument("--input", type=Path,
                        help="a JSON array of ruleset objects in GitHub REST format")
    parser.add_argument("--export", action="store_true",
                        help="write the live rulesets to .github/rulesets/ instead of comparing")
    parser.add_argument("--summary", type=Path, help="append the Markdown report to this file")
    args = parser.parse_args(argv)
    if args.repository and not re.fullmatch(r"[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+", args.repository):
        parser.error("repository must have the form owner/name")

    root = args.root.resolve()
    try:
        if args.repository:
            live = load_live(args.repository, os.environ.get("GITHUB_TOKEN", ""))
        else:
            data = json.loads(args.input.read_text(encoding="utf-8"))
            live = {r["name"]: r for r in (normalize(item) for item in data)}
        if args.export:
            directory = root / RULESET_DIR
            directory.mkdir(parents=True, exist_ok=True)
            for stale in directory.glob("*.json"):
                stale.unlink()
            for ruleset in live.values():
                (directory / file_name(ruleset["name"])).write_text(render(ruleset),
                                                                    encoding="utf-8")
            print(f"wrote {len(live)} ruleset(s) to {RULESET_DIR.as_posix()}/")
            return 0
        findings = compare(load_versioned(root), live)
    except (OSError, ValueError, KeyError, TypeError, urllib.error.URLError) as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 2

    status = "FAIL" if findings else "PASS"
    lines = ["# Ruleset drift", "", f"Status: **{status}** · live rulesets: {len(live)}", ""]
    lines += [f"- {finding}" for finding in findings] or ["Every live ruleset matches its versioned copy."]
    report = "\n".join(lines) + "\n"
    print(report, end="")
    if args.summary:
        with args.summary.open("a", encoding="utf-8") as handle:
            handle.write(report)
    return 1 if findings else 0


if __name__ == "__main__":
    raise SystemExit(main())
