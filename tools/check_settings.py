#!/usr/bin/env python3
"""Compare the repository's live settings with .github/repository-settings.json.

Rulesets are versioned under .github/rulesets/ and checked by
tools/check_rulesets.py. The other settings that protect the repository live in
the GitHub settings pages: the merge methods and repository features, secret
scanning and the other security features, and the GitHub Actions policy. Their
intended values are versioned in .github/repository-settings.json, and this
check reports, for each one, whether the live value matches, differs, or could
not be read.

Several of these settings are visible only to repository administrators. A
setting whose endpoint answers 401, 403, 404 or 409, or whose field is absent
from the response, is reported as not verifiable: neither a match nor a
difference. Run the check with an administrator token to verify every setting.

Read-only: it never changes a setting. To change one, change it in the GitHub
settings, update the file in a pull request that explains why, and run this
check with an administrator token.

Exit status: 0 when no setting differs, 1 when one differs (or, with --strict,
when one is not verifiable), 2 on a usage, network or repository error.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
import urllib.error
import urllib.request
from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
SETTINGS_FILE = Path(".github") / "repository-settings.json"
API = "https://api.github.com"

# Statuses that mean "this token may not read it" rather than an error. The
# selected-actions endpoint answers 409 when the policy is not "selected".
UNREADABLE = (401, 403, 404, 409)

# Each recorded or live response is identified by its path below
# /repos/{owner}/{name}; "repository" is the repository object itself.
ENDPOINTS = {
    "repository": "",
    "private-vulnerability-reporting": "/private-vulnerability-reporting",
    "vulnerability-alerts": "/vulnerability-alerts",
    "actions/permissions": "/actions/permissions",
    "actions/permissions/selected-actions": "/actions/permissions/selected-actions",
    "actions/permissions/workflow": "/actions/permissions/workflow",
}

GENERAL = (
    "allow_auto_merge", "allow_forking", "allow_merge_commit", "allow_rebase_merge",
    "allow_squash_merge", "allow_update_branch", "delete_branch_on_merge",
    "has_discussions", "has_issues", "has_projects", "has_wiki", "visibility",
    "web_commit_signoff_required",
)
SECURITY_STATUSES = (
    "dependabot_security_updates", "secret_scanning", "secret_scanning_push_protection",
)
SELECTED_ACTIONS = ("github_owned_allowed", "patterns_allowed", "verified_allowed")


@dataclass(frozen=True)
class Response:
    """One answer from an endpoint: its status and its JSON body, if any."""

    status: int
    body: Any


@dataclass(frozen=True)
class Result:
    """The outcome for one setting: "match", "differ" or "unverifiable"."""

    name: str
    outcome: str
    expected: Any
    detail: Any


class Unreadable(Exception):
    """The live value of a setting cannot be read with the available access."""


Get = Callable[[str], Response]
Reader = Callable[[Get], Any]


def body(get: Get, endpoint: str) -> Any:
    response = get(endpoint)
    if response.status in UNREADABLE:
        raise Unreadable(f"HTTP {response.status} from {endpoint}")
    return response.body


def field(data: Any, endpoint: str, *path: str) -> Any:
    for key in path:
        if not isinstance(data, dict) or key not in data:
            raise Unreadable(f"{endpoint} has no {'.'.join(path)} field")
        data = data[key]
    return data


def flag(value: Any, endpoint: str) -> str:
    if not isinstance(value, bool):
        raise ValueError(f"{endpoint} returned a non-boolean enabled value")
    return "enabled" if value else "disabled"


def is_admin(get: Get) -> bool:
    try:
        return field(body(get, "repository"), "repository", "permissions", "admin") is True
    except Unreadable:
        return False


def read_dependabot_alerts(get: Get) -> str:
    # 204 means enabled. 404 means disabled only to an administrator; to
    # anyone else it is also the answer for "you may not read this".
    status = get("vulnerability-alerts").status
    if status == 204:
        return "enabled"
    if status == 404 and is_admin(get):
        return "disabled"
    if status in UNREADABLE:
        raise Unreadable(f"HTTP {status} from vulnerability-alerts"
                         + (" without administrator access" if status == 404 else ""))
    raise ValueError(f"vulnerability-alerts answered HTTP {status}")


def read_selected_actions(get: Get) -> dict[str, Any]:
    data = body(get, "actions/permissions/selected-actions")
    selected = {key: field(data, "actions/permissions/selected-actions", key)
                for key in SELECTED_ACTIONS}
    selected["patterns_allowed"] = sorted(selected["patterns_allowed"])
    return selected


def reader(endpoint: str, *path: str) -> Reader:
    return lambda get: field(body(get, endpoint), endpoint, *path)


READERS: dict[str, dict[str, Reader]] = {
    "general": {key: reader("repository", key) for key in GENERAL},
    "security_and_analysis": {
        **{key: reader("repository", "security_and_analysis", key, "status")
           for key in SECURITY_STATUSES},
        "dependabot_alerts": read_dependabot_alerts,
        "private_vulnerability_reporting": lambda get: flag(
            reader("private-vulnerability-reporting", "enabled")(get),
            "private-vulnerability-reporting"),
    },
    "actions": {
        "enabled": reader("actions/permissions", "enabled"),
        "allowed_actions": reader("actions/permissions", "allowed_actions"),
        "selected_actions": read_selected_actions,
        "default_workflow_permissions": reader("actions/permissions/workflow",
                                               "default_workflow_permissions"),
        "can_approve_pull_request_reviews": reader("actions/permissions/workflow",
                                                   "can_approve_pull_request_reviews"),
    },
}


def render(settings: dict[str, Any]) -> str:
    return json.dumps(settings, indent=2, sort_keys=True, ensure_ascii=False) + "\n"


def load_expected(root: Path) -> dict[str, dict[str, Any]]:
    """Read the versioned settings and reject unknown or unnormalized entries."""
    path = root / SETTINGS_FILE
    text = path.read_text(encoding="utf-8")
    data = json.loads(text)
    if not isinstance(data, dict):
        raise ValueError(f"{SETTINGS_FILE.as_posix()} is not a JSON object")
    for group, settings in data.items():
        if group not in READERS:
            raise ValueError(f"{SETTINGS_FILE.as_posix()}: unknown group {group!r}")
        if not isinstance(settings, dict):
            raise ValueError(f"{SETTINGS_FILE.as_posix()}: {group} is not a JSON object")
        unknown = sorted(set(settings) - set(READERS[group]))
        if unknown:
            raise ValueError(f"{SETTINGS_FILE.as_posix()}: unknown {group} setting(s) "
                             f"{', '.join(unknown)}")
    if text != render(data):
        raise ValueError(f"{SETTINGS_FILE.as_posix()} is not normalized: use two-space "
                         "indentation, sorted keys and a final newline")
    return data


def compare(expected: dict[str, dict[str, Any]], get: Get) -> list[Result]:
    results = []
    for group in sorted(expected):
        for key in sorted(expected[group]):
            name = f"{group}.{key}"
            want = expected[group][key]
            try:
                live = READERS[group][key](get)
            except Unreadable as reason:
                results.append(Result(name, "unverifiable", want, str(reason)))
                continue
            results.append(Result(name, "match" if live == want else "differ", want, live))
    return results


def cached(fetch: Get) -> Get:
    responses: dict[str, Response] = {}

    def get(endpoint: str) -> Response:
        if endpoint not in responses:
            response = fetch(endpoint)
            if response.status not in (200, 204, *UNREADABLE):
                raise ValueError(f"{endpoint} answered HTTP {response.status}")
            responses[endpoint] = response
        return responses[endpoint]
    return get


def live_fetcher(repository: str, token: str) -> Get:
    headers = {"Accept": "application/vnd.github+json", "X-GitHub-Api-Version": "2022-11-28"}
    if token:
        headers["Authorization"] = f"Bearer {token}"

    def fetch(endpoint: str) -> Response:
        url = f"{API}/repos/{repository}{ENDPOINTS[endpoint]}"
        request = urllib.request.Request(url, headers=headers)
        try:
            with urllib.request.urlopen(request, timeout=30) as response:
                raw = response.read()
                return Response(response.status, json.loads(raw) if raw else None)
        except urllib.error.HTTPError as error:
            if error.code in UNREADABLE:
                return Response(error.code, None)
            raise
    return fetch


def recorded_fetcher(path: Path) -> Get:
    """Serve responses recorded as {endpoint: {"status": int, "body": ...}}."""
    recorded = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(recorded, dict):
        raise ValueError(f"{path} is not a JSON object")

    def fetch(endpoint: str) -> Response:
        if endpoint not in recorded:
            raise ValueError(f"{path} has no recorded response for {endpoint}")
        entry = recorded[endpoint]
        return Response(int(entry["status"]), entry.get("body"))
    return fetch


def show(value: Any) -> str:
    return f"`{json.dumps(value, sort_keys=True)}`"


def report(results: list[Result], strict: bool) -> tuple[str, int]:
    groups = {outcome: [r for r in results if r.outcome == outcome]
              for outcome in ("differ", "unverifiable", "match")}
    failed = bool(groups["differ"]) or (strict and bool(groups["unverifiable"]))
    status = "FAIL" if failed else "INCOMPLETE" if groups["unverifiable"] else "PASS"
    lines = ["# Repository settings drift", "",
             f"Status: **{status}** · differ: {len(groups['differ'])} · "
             f"not verifiable: {len(groups['unverifiable'])} · match: {len(groups['match'])}",
             ""]
    sections: tuple[tuple[str, str, Callable[[Result], str]], ...] = (
        ("differ", "Differences",
         lambda r: f"- `{r.name}`: expected {show(r.expected)}, live {show(r.detail)}"),
        ("unverifiable", "Not verifiable",
         lambda r: f"- `{r.name}`: {r.detail}; expected {show(r.expected)}"),
        ("match", "Matches", lambda r: f"- `{r.name}`: {show(r.expected)}"),
    )
    for outcome, title, line in sections:
        if groups[outcome]:
            lines += [f"## {title}", "", *(line(r) for r in groups[outcome]), ""]
    if groups["unverifiable"]:
        lines.append("Settings that are not verifiable need an administrator token; "
                     + ("--strict counts them as failures." if strict
                        else "they do not change the exit status without --strict."))
    return "\n".join(lines).rstrip("\n") + "\n", 1 if failed else 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0] if __doc__ else None)
    parser.add_argument("--root", type=Path, default=ROOT)
    source = parser.add_mutually_exclusive_group(required=True)
    source.add_argument("--repository", help="owner/name; reads GITHUB_TOKEN if set")
    source.add_argument("--input", type=Path,
                        help='recorded responses: a JSON object {endpoint: {"status", "body"}}')
    parser.add_argument("--strict", action="store_true",
                        help="fail when a setting is not verifiable")
    parser.add_argument("--summary", type=Path, help="append the Markdown report to this file")
    args = parser.parse_args(argv)
    if args.repository and not re.fullmatch(r"[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+", args.repository):
        parser.error("repository must have the form owner/name")

    try:
        expected = load_expected(args.root.resolve())
        fetch = (live_fetcher(args.repository, os.environ.get("GITHUB_TOKEN", ""))
                 if args.repository else recorded_fetcher(args.input))
        results = compare(expected, cached(fetch))
    except (OSError, ValueError, KeyError, TypeError, urllib.error.URLError) as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 2

    text, status = report(results, args.strict)
    print(text, end="")
    if args.summary:
        with args.summary.open("a", encoding="utf-8") as handle:
            handle.write(text)
    return status


if __name__ == "__main__":
    raise SystemExit(main())
