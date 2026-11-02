#!/usr/bin/env python3
"""Check issue metadata and dependency status that issue forms cannot enforce.

Read-only: the check never changes an issue. It reports findings for me to
resolve by hand (docs/ISSUE_LABELS.md).
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import os
import re
import subprocess
import sys
import urllib.error
import urllib.request
from collections.abc import Collection, Iterable, Mapping, Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import Any

PRIORITIES = ("P1", "P2", "P3")
BLOCKED = "blocked"
FORM_PRIORITIES = {"bug.yml": "P2", "enhancement.yml": "P3", "documentation.yml": "P3",
                   "security-contact.yml": "P2"}
# Forms whose reporter proposes the target end date; triage adds it to the others.
TARGET_DATE_FORMS = ("bug.yml", "enhancement.yml", "documentation.yml")
# The planning line triage writes, or the field an issue form renders. Each
# captures its whole value, so an invalid line never falls back to the form.
PLANNED_TARGET = re.compile(r"^[ \t]*[-*][ \t]*Target end date:[ \t]*(.*?)[ \t]*$", re.I | re.M)
FORM_TARGET = re.compile(
    r"^###\s+(?:🗓️\s*)?Target end date(?:\s*\(Europe/Rome\))?\s*$\n\s*(.*?)[ \t]*$", re.I | re.M)
ROME_SUFFIX = re.compile(r"\s*\(Europe/Rome\)$", re.I)
DEPENDENCY_STATUS_WORDS = re.compile(r"block|ready|waiting|depend|on[ -]hold|prerequisite", re.I)
HEADING = re.compile(r"^#{2,3}\s+(.*)$")
EXTERNAL = re.compile(r"^\s*(?:[-*]\s*)?external condition:\s*\S", re.I)
API = "https://api.github.com"


@dataclass(frozen=True)
class Issue:
    number: int
    state: str  # "open" or "closed"
    labels: tuple[str, ...]
    assignees: tuple[str, ...]
    milestone: str | None
    body: str
    # Calendar date (YYYY-MM-DD) of the milestone's due_on, as GitHub stores it.
    milestone_due: str | None = None
    milestone_state: str = "open"


@dataclass(frozen=True)
class Finding:
    number: int
    rule: str
    message: str


def normalized(body: str) -> str:
    """Use LF line endings: issues edited in a browser are saved with CRLF."""
    return body.replace("\r\n", "\n").replace("\r", "\n")


def issue_from_api(item: Mapping[str, Any]) -> Issue:
    milestone = item.get("milestone")
    return Issue(
        number=int(item["number"]),
        state=str(item["state"]).lower(),
        labels=tuple(str(label["name"]) for label in item.get("labels", [])),
        assignees=tuple(str(user["login"]) for user in item.get("assignees", [])),
        milestone=str(milestone["title"]) if milestone else None,
        body=normalized(str(item.get("body") or "")),
        milestone_due=str(milestone["due_on"])[:10] if milestone and milestone.get("due_on") else None,
        milestone_state=str(milestone.get("state") or "open").lower() if milestone else "open",
    )


def dependency_section(body: str) -> list[str] | None:
    """Return the lines of the issue's dependency section, or None if it has none."""
    lines = body.splitlines()
    for index, line in enumerate(lines):
        match = HEADING.match(line)
        if match and "dependencies" in match.group(1).lower():
            section = []
            for following in lines[index + 1:]:
                if HEADING.match(following):
                    break
                section.append(following)
            return section
    return None


def references(lines: Iterable[str], repository: str, own: int) -> set[int]:
    url = re.compile(rf"github\.com/{re.escape(repository)}/issues/(\d+)", re.I)
    found: set[int] = set()
    for line in lines:
        found.update(int(n) for n in re.findall(r"(?<![\w/&])#(\d+)\b", line))
        found.update(int(n) for n in url.findall(line))
    found.discard(own)
    return found


def target_end_date(body: str) -> str | None:
    """Read the planning line or the issue-form field without guessing a date."""
    body = normalized(body)
    match = PLANNED_TARGET.search(body) or FORM_TARGET.search(body)
    return ROME_SUFFIX.sub("", match.group(1)) if match else None


def iso_date(text: str | None) -> dt.date | None:
    if text is None or not re.fullmatch(r"\d{4}-\d{2}-\d{2}", text):
        return None
    try:
        return dt.date.fromisoformat(text)
    except ValueError:
        return None


def target_date_findings(issue: Issue) -> list[Finding]:
    """The target end date must be a real ISO date before the milestone's due date."""
    target = iso_date(target_end_date(issue.body))
    if target is None:
        return [Finding(issue.number, "target-date",
                        "needs a valid ISO YYYY-MM-DD target end date (Europe/Rome)")]
    if not issue.milestone:
        return []  # already reported by the milestone rule
    due = iso_date(issue.milestone_due)
    if due is None:
        return [Finding(issue.number, "milestone-due",
                        f"milestone {issue.milestone!r} has no due date, so the target end "
                        "date cannot be checked against it")]
    if target >= due:
        return [Finding(issue.number, "milestone-due",
                        f"target end date {target} must be before the milestone due date {due}")]
    return []


def metadata_findings(issue: Issue, assignee: str) -> list[Finding]:
    findings = []
    if assignee not in issue.assignees:
        findings.append(Finding(issue.number, "assignee", f"is not assigned to {assignee}"))
    priorities = [label for label in issue.labels if label in PRIORITIES]
    if len(priorities) != 1:
        found = ", ".join(priorities) or "none"
        findings.append(Finding(issue.number, "priority",
                                f"needs exactly one of P1/P2/P3 (found: {found})"))
    if "P1" in priorities and "bug" not in issue.labels:
        findings.append(Finding(issue.number, "p1-reserved",
                                "has P1 without bug: P1 is reserved for confirmed release-blocking defects"))
    if not issue.milestone:
        findings.append(Finding(issue.number, "milestone", "has no milestone"))
    elif issue.milestone_state != "open":
        findings.append(Finding(issue.number, "milestone",
                                f"is in the closed milestone {issue.milestone!r}: "
                                "move it to an open milestone"))
    return findings


def dependency_findings(issue: Issue, states: Mapping[int, str], repository: str,
                        pulls: Collection[int] = ()) -> list[Finding]:
    """A named pull request is a known reference, never a prerequisite."""
    section = dependency_section(issue.body)
    if section is None:
        return [Finding(issue.number, "dependencies", "has no dependency section")]
    named = references(section, repository, issue.number)
    missing = sorted(n for n in named if n not in states and n not in pulls)
    open_prerequisites = sorted(n for n in named if states.get(n) == "open")
    external = any(EXTERNAL.match(line) for line in section)
    blocked = BLOCKED in issue.labels
    findings = [Finding(issue.number, "dependencies", f"names #{n}, which does not exist")
                for n in missing]
    if blocked and not open_prerequisites and not external:
        findings.append(Finding(issue.number, "blocked",
                                "is blocked but names no open prerequisite or external "
                                "condition: remove blocked or name the blocker"))
    if not blocked and open_prerequisites:
        listed = ", ".join(f"#{n}" for n in open_prerequisites)
        findings.append(Finding(issue.number, "blocked",
                                f"names open prerequisite(s) {listed} but is not blocked"))
    if not blocked and external:
        findings.append(Finding(issue.number, "blocked",
                                "names an external condition but is not blocked"))
    return findings


def audit(issues: Sequence[Issue], states: Mapping[int, str], repository: str,
          catalogue: set[str] | None = None, pulls: Collection[int] = (),
          *, assignee: str) -> list[Finding]:
    """Audit every open issue; ``states`` maps every known issue number to its state
    and ``pulls`` holds the numbers of the repository's pull requests."""
    findings = []
    for issue in sorted(issues, key=lambda item: item.number):
        if issue.state != "open":
            continue
        findings += metadata_findings(issue, assignee)
        findings += target_date_findings(issue)
        findings += dependency_findings(issue, states, repository, pulls)
        if catalogue is not None:
            findings += [Finding(issue.number, "catalogue", f"has label {label!r}, "
                                 "which is not in .github/labels.json")
                         for label in issue.labels if label not in catalogue]
    return findings


def catalogue_labels(root: Path) -> set[str]:
    # Use the same schema validator as synchronization; never silently ignore overlays.
    result = subprocess.run(
        ["node", str(root / ".github/scripts/labels-sync.mjs"), "--mode", "validate",
         "--manifest", str(root / ".github/labels.json")],
        capture_output=True, text=True, encoding="utf-8", errors="replace", check=False,
    )
    if result.returncode:
        raise ValueError(result.stderr.strip() or result.stdout.strip())
    data = json.loads((root / ".github" / "labels.json").read_text(encoding="utf-8"))
    return {str(label["name"]) for label in data["labels"]}


def configured_assignee(root: Path) -> str:
    """Read the required person from the issue form, not the repository owner."""
    text = (root / ".github/ISSUE_TEMPLATE/bug.yml").read_text(encoding="utf-8")
    assignees = flow_list(text, "assignees")
    if len(assignees) != 1:
        raise ValueError("bug.yml must configure exactly one assignee")
    return assignees[0]


def config_findings(root: Path) -> list[str]:
    """Check the label catalogue and the issue forms against the issue rules."""
    problems = []
    labels = catalogue_labels(root)
    assignee = configured_assignee(root)
    status_labels = sorted(n for n in labels if DEPENDENCY_STATUS_WORDS.search(n))
    if status_labels != [BLOCKED]:
        problems.append(f"the dependency status labels must be exactly ['blocked'], "
                        f"found {status_labels}")
    forms = root / ".github" / "ISSUE_TEMPLATE"
    for name, priority in sorted(FORM_PRIORITIES.items()):
        path = forms / name
        if not path.is_file():
            problems.append(f"{path.relative_to(root)} is missing")
            continue
        text = path.read_text(encoding="utf-8")
        form_labels = flow_list(text, "labels")
        if [label for label in form_labels if label in PRIORITIES] != [priority]:
            problems.append(f"{name} must set exactly the priority {priority}")
        problems += [f"{name} sets unknown label {label!r}"
                     for label in form_labels if label not in labels]
        if flow_list(text, "assignees") != [assignee]:
            problems.append(f"{name} must assign {assignee}")
        if not re.search(r"^\s*id:\s*dependencies\s*$", text, re.M):
            problems.append(f"{name} has no dependencies field")
        if name in TARGET_DATE_FORMS and not re.search(r"^\s*id:\s*target_end_date\s*$",
                                                       text, re.M):
            problems.append(f"{name} has no target end date field")
    config = forms / "config.yml"
    if not config.is_file() or not re.search(r"^blank_issues_enabled:\s*false\s*$",
                                             config.read_text(encoding="utf-8"), re.M):
        problems.append("ISSUE_TEMPLATE/config.yml must set blank_issues_enabled: false")
    return problems


def flow_list(text: str, key: str) -> list[str]:
    match = re.search(rf"^{key}:\s*\[(.*)\]\s*$", text, re.M)
    if not match:
        return []
    return [item.strip().strip("\"'") for item in match.group(1).split(",") if item.strip()]


def fetch(url: str, token: str) -> Any:
    request = urllib.request.Request(url, headers={
        "Accept": "application/vnd.github+json",
        "Authorization": f"Bearer {token}",
        "X-GitHub-Api-Version": "2022-11-28",
    })
    with urllib.request.urlopen(request, timeout=30) as response:
        return json.load(response)


Loaded = tuple[list[Issue], dict[int, str], set[int]]


def split(items: Iterable[Mapping[str, Any]]) -> Loaded:
    """Return the issues, their states and the pull request numbers of REST items."""
    issues: list[Issue] = []
    pulls: set[int] = set()
    for item in items:
        if "pull_request" in item:
            pulls.add(int(item["number"]))
        else:
            issues.append(issue_from_api(item))
    return issues, {issue.number: issue.state for issue in issues}, pulls


def load_live(repository: str, token: str) -> Loaded:
    """Fetch every issue and pull request (open and closed) of the repository."""
    items: list[Mapping[str, Any]] = []
    for page in range(1, 101):
        batch = fetch(f"{API}/repos/{repository}/issues?state=all&per_page=100&page={page}",
                      token)
        items += batch
        if len(batch) < 100:
            break
    else:
        raise RuntimeError("10000 or more issues; refusing to audit a possibly partial list")
    return split(items)


def load_file(path: Path) -> Loaded:
    return split(json.loads(path.read_text(encoding="utf-8")))


def render(findings: Sequence[Finding], audited: int) -> str:
    status = "PASS" if not findings else "FAIL"
    lines = ["# Issue triage", "", f"Status: **{status}**; open issues audited: {audited}", ""]
    if findings:
        lines += ["| Issue | Rule | Finding |", "| --- | --- | --- |"]
        lines += [f"| #{f.number} | `{f.rule}` | {f.message} |" for f in findings]
        lines += ["", "Resolve each finding by hand; this check never changes an issue. "
                  "See docs/ISSUE_LABELS.md."]
    else:
        lines.append("Every open issue meets the triage rules.")
    return "\n".join(lines) + "\n"


def parse_args(argv: Sequence[str] | None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--root", type=Path, default=Path(__file__).resolve().parents[1])
    source = parser.add_mutually_exclusive_group(required=True)
    source.add_argument("--repository", help="owner/name; reads GITHUB_TOKEN")
    source.add_argument("--input", type=Path, help="issues exported as GitHub REST JSON")
    source.add_argument("--check-config", action="store_true",
                        help="check the issue forms and label catalogue only (offline)")
    parser.add_argument("--repo-name",
                        help="repository used to resolve issue URLs with --input")
    parser.add_argument("--output", type=Path, help="write the findings as JSON")
    parser.add_argument("--summary", type=Path, help="append the Markdown report to this file")
    args = parser.parse_args(argv)
    if args.input and not args.repo_name:
        parser.error("--input requires --repo-name owner/name")
    repository = args.repository or args.repo_name
    if repository and not re.fullmatch(r"[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+", repository):
        parser.error("repository must have the form owner/name")
    return args


def main(argv: Sequence[str] | None = None) -> int:
    args = parse_args(argv)
    root = args.root.resolve()
    try:
        if args.check_config:
            problems = config_findings(root)
            for problem in problems:
                print(f"FAIL: {problem}")
            print(f"issue-config: {'PASS' if not problems else 'FAIL'}")
            return 1 if problems else 0
        if args.repository:
            token = os.environ.get("GITHUB_TOKEN", "")
            if not token:
                raise RuntimeError("--repository requires GITHUB_TOKEN")
            issues, states, pulls = load_live(args.repository, token)
            repository = args.repository
        else:
            issues, states, pulls = load_file(args.input)
            repository = args.repo_name
        findings = audit(issues, states, repository, catalogue_labels(root), pulls,
                         assignee=configured_assignee(root))
    except (OSError, RuntimeError, ValueError, KeyError, urllib.error.URLError) as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 2
    report = render(findings, sum(1 for issue in issues if issue.state == "open"))
    print(report, end="")
    if args.summary:
        with args.summary.open("a", encoding="utf-8") as handle:
            handle.write(report)
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(json.dumps(
            {"status": "pass" if not findings else "fail",
             "findings": [vars(f) for f in findings]}, indent=2) + "\n", encoding="utf-8")
    return 1 if findings else 0


if __name__ == "__main__":
    raise SystemExit(main())
