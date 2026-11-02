# 🏷️ Issue labels & automation

![Labels](https://img.shields.io/badge/labels-27-1D76DB)

This repository defines 27 issue labels, each with a name, description, and colour. `.github/labels.json` is authoritative; edit it through a pull request. `prune: true` removes labels outside the catalogue, including their issue associations, so rename a live label on GitHub before renaming it in the catalogue.

## ⚙️ Automation

| Workflow | Purpose |
| --- | --- |
| Sync issue labels | Validates pull requests; reconciles labels after a catalogue change lands on the default branch or a manual run on the default branch. Verifies the exact live result. |
| Detect issue-label drift | Checks daily at 05:17 UTC and manually; PR runs use offline fixtures. Reports drift without modifying labels. |
| Issue triage | Audits open issues on issue and milestone events or manual runs. Reports a missing assignee, priority, open milestone, target end date or dependency section, a target end date not before the milestone's due date, invalid labels, and inconsistent blocked status. |
| Keep milestones issue-only | Removes any milestone set on a pull request when the pull request is opened, reopened, edited, or given a milestone, and in a daily and manual scan of every pull request. It never checks out pull-request code. |
| Static checks | Runs the label scripts' self-tests, validates the issue forms against the catalogue (`python3 tools/check_issues.py --check-config`), and runs the triage unit tests (`python3 -m unittest discover -s tests`). |

Among these workflows, only synchronization has `issues: write` and only *Keep milestones issue-only* has `pull-requests: write`. PR validation is read-only. The label and triage workflows pin actions to full commit SHAs.

My GitHub account, `danielep71`, is the assignee in the four issue forms (`assignees`); the triage check reads it from `bug.yml` and fails if the forms disagree.

## 🧭 Using the labels

Every issue has exactly one priority, the type labels that fit the work, and the area labels it touches. Pull requests carry the labels of the work they deliver; milestones belong to issues only, and *Keep milestones issue-only* removes any milestone set on a pull request.

### Priorities

| Label | Use |
| --- | --- |
| `P1` | Reserved for a confirmed release-blocking correctness or safety defect, such as a failure that leaves Excel unrecoverable or applies state to the wrong window. Only together with `bug`; never for setup, documentation, or tooling work. |
| `P2` | Material correctness, robustness, or maintainability defect. Default of the bug and private security contact forms. |
| `P3` | Non-blocking quality, testing, documentation, setup, or technical-debt work. Default of the enhancement and documentation forms. |

A form sets the default priority; triage confirms or changes it. Raise a bug to `P1` only once the defect is confirmed.

Every open issue also needs a target end date: an ISO `YYYY-MM-DD` calendar date (Europe/Rome) earlier than the due date of its milestone. The bug, enhancement, and documentation forms ask for it in the *Target end date* field. On other issues, including private security contacts, add it during triage as a line in the issue body:

```text
- Target end date: YYYY-MM-DD (Europe/Rome)
```

A note may follow the time zone after a semicolon, as in `(Europe/Rome; same as Target date)`. When both are present, that line takes precedence over the form field, so edit the line to reschedule. Triage reports a missing or invalid date (`target-date`), and a milestone without a due date or a target end date on or after it (`milestone-due`).

Every open issue also needs a section whose heading contains "Dependencies". Name each prerequisite there as an issue link (`#12`) or on a line starting `External condition:`, or write "None".

### Types

| Label | Use |
| --- | --- |
| `bug` | A confirmed defect. The bug form sets it. |
| `enhancement` | New capability or improvement. The enhancement form sets it. |
| `documentation` | Documentation work. The documentation form sets it. |
| `behavior-change` | Add to work that changes externally observable behavior, such as `UI_*` results, diagnostics, or restored state. |
| `refactor` | Restructuring that preserves supported behavior. If supported behavior changes, use `behavior-change` instead. |
| `tests` | Regression coverage, test harnesses, fixtures, or manual validation. |
| `ci` | Workflows, automated checks, and dependency updates. Release certification belongs to `release`. |
| `repository` | Repository structure, governance, metadata, or tooling. |
| `release` | Release preparation, certification, evidence, packaging, or publication. |
| `security` | Hardening and supply-chain work, and private security contact requests, which never contain vulnerability details. Never for an undisclosed vulnerability: report it as `SECURITY.md` describes. |

### Areas

| Label | Use |
| --- | --- |
| `api` | The supported `UI_*` facade: signatures, parameters, enums, results, or compatibility. |
| `runtime` | Shared runtime services, diagnostics, result buffers, or host-state management. |
| `ribbon` | Ribbon visibility, restoration, or Ribbon-specific policy. |
| `titlebar` | WinAPI title-bar and frame state, ownership, identity, or recovery. |
| `snapshots` | UI snapshot capture, identity, lifecycle, or restoration. |
| `sdi` | Excel SDI multi-window identity, targeting, or active-window behavior. |
| `demo` | Demo VBA source, demo builder, or distributed demo workbook. |

### Status and triage

| Label | Use |
| --- | --- |
| `blocked` | A prerequisite named in the dependency section, as an issue link or an `External condition:` line, is unresolved. Remove it once all are resolved; ready issues carry no status label. |
| `question` | I need information or a decision before the work can proceed. |
| `help wanted` | I welcome outside investigation, implementation, validation, or review. |
| `good first issue` | Well-scoped task suitable for a first contribution. |
| `duplicate` | Close the issue and link the issue or pull request that already covers the work. |
| `invalid` | Close the issue: it is incomplete, unsupported, or not actionable as submitted. |
| `wontfix` | Close the issue: I have decided not to do this work. |

## ✅ Closing issues

Close an issue as completed only when linked evidence meets its acceptance criteria, and check an acceptance box only when that evidence demonstrates it. When the issue is closed, add a **Resolution** section to its body with:

- links to the change (pull request and commit);
- the checks actually run, with their results;
- what was not validated, and any accepted limitation.

Close for another reason (duplicate, not planned, or superseded) only with an explicit explanation; the `duplicate`, `invalid`, and `wontfix` rows above say which label applies. Record material scope or policy decisions in the issue, pull request, or commit description.

## 🎨 Catalogue

The table below lists the descriptions defined in `.github/labels.json`, which synchronization uses as its source of truth. `python3 tools/check_repo.py` fails if this table, the label count, or the badge disagree with it.

| Label | Colour | Description |
| --- | --- | --- |
| `api` | `#0052CC` | Supported VBA API, signatures, parameters, enums, or compatibility contract |
| `behavior-change` | `#C24E00` | Intentional externally observable behavior change |
| `blocked` | `#B60205` | Cannot proceed until a named dependency is complete |
| `bug` | `#D73A4A` | Confirmed defect in behavior, correctness, error handling, or a documented contract |
| `ci` | `#24292F` | Automated quality checks, workflow maintenance, or dependency updates |
| `demo` | `#F9D0C4` | Demo VBA source, demo builder, or distributed demo workbook |
| `documentation` | `#0075CA` | Documentation for users or for my repository maintenance |
| `duplicate` | `#CFD3D7` | An existing issue or pull request already covers this work |
| `enhancement` | `#A2EEEF` | New capability or meaningful improvement to supported behavior |
| `good first issue` | `#7057FF` | Well-scoped task suitable for a first contribution |
| `help wanted` | `#008672` | I welcome external investigation, implementation, validation, or review |
| `invalid` | `#E4E669` | Report is incomplete, unsupported, or not actionable as submitted |
| `P1` | `#B60205` | Confirmed release-blocking correctness or safety defect (bugs only) |
| `P2` | `#D93F0B` | Material correctness, robustness, or maintainability defect |
| `P3` | `#FBCA04` | Non-blocking quality, testing, documentation, setup or technical-debt work |
| `question` | `#D876E3` | I need further information or need to make a decision |
| `refactor` | `#6F42C1` | Internal restructuring intended to preserve supported behavior |
| `release` | `#B60276` | Release preparation, certification, evidence, packaging, or publication |
| `repository` | `#1D76DB` | Repository structure, governance, metadata, tooling, or maintenance |
| `ribbon` | `#D4C5F9` | Ribbon visibility, restoration, or Ribbon-specific policy |
| `runtime` | `#2B7489` | Shared runtime services, diagnostics, result buffers, or host-state management |
| `sdi` | `#BFD4F2` | Excel SDI multi-window identity, targeting, or active-window behavior |
| `security` | `#8B0000` | Security boundary, vulnerability handling, supply chain, or hardening |
| `snapshots` | `#0E8A16` | UI snapshot capture, identity, lifecycle, or restoration |
| `tests` | `#006B75` | Regression coverage, test harnesses, fixtures, or manual validation |
| `titlebar` | `#5319E7` | WinAPI title-bar and frame state, ownership, identity, or recovery |
| `wontfix` | `#6E7781` | I have decided not to implement or continue this work |

## 🔍 Local issue audit

The audit needs Python 3.11+ and Node.js, because it validates the catalogue with the synchronization script. To audit the live issues with your GitHub CLI login:

```sh
GITHUB_TOKEN="$(gh auth token)" python3 tools/check_issues.py --repository danielep71/VBA-EXCEL_UI
```

To audit an export instead, pass a JSON array of issues and pull requests, open and closed, in GitHub REST format, and name the repository explicitly so dependency URLs resolve:

```sh
python3 tools/check_issues.py --input issues.json --repo-name danielep71/VBA-EXCEL_UI
```

Triage accepts only the labels defined in `.github/labels.json`.
