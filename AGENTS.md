# Repository instructions

## Documentation voice

- Write my decisions and responsibilities in first person: “I”, “me”, and “my”.
- Do not describe me in third person as “the maintainer” or “the owner”.
- Address instructions for contributors and users directly to the reader.

## Repository scope

- Describe this repository's current files, configuration, and behavior.
- Do not insert references to other repositories, copied-from history, unrelated projects, or conversation context into operational documentation.
- Keep necessary platform/dependency documentation links and third-party license notices. Put licensing attribution in dedicated license files, not operational guides.
- Do not invent project history, capabilities, or dependencies.

## Branch names

Use descriptive prefixes such as `docs/`, `fix/`, `feat/`, `test/`, or `chore/`. Do not use `codex/` or `claude/`.

## Commit authorship

Author every commit as me, not under an assistant's identity. Before committing, set the repository's Git `user.name` and `user.email` to the name and GitHub noreply address (`<id>+<login>@users.noreply.github.com`) that my own commits in this repository use. Do not add a `Co-Authored-By` trailer that credits an assistant. Never rewrite existing history to change authors.

## Audience separation

Keep editorial rules, assistant instructions, branch-naming preferences, and collaboration conventions in this file. Reader-facing documentation must explain the repository and its usage, not our writing process. Apply the requested voice naturally; do not explain that prose is written “in my voice”.

## Review findings

When a review of this repository finds something to fix, record each finding, or each group of findings that one change will fix, as an issue before or instead of fixing it silently. Each issue must pass the triage rules in `docs/ISSUE_LABELS.md`; in short, it must be:

- assigned to me;
- labelled with exactly one priority label, plus the type and area labels that fit;
- placed in the open milestone with the earliest due date;
- given a `- Target end date: YYYY-MM-DD (Europe/Rome)` line earlier than that milestone's due date;
- given a dependency section that names prerequisites or says "None".

Describe the finding with its file and line, the evidence, why it matters, and the proposed fix, so the issue can be resolved without the review that found it.

## Validation

Before pushing, run `python3 tools/verify.py --ci` and fix every finding. It runs every static check in one declared order (`--list` shows them) and needs Node.js, plus Ruff and mypy installed from `tools/requirements-quality-ci.txt`. Fix formatting drift with `python3 tools/reformat.py --write src/*.bas test/*.bas demo/*.bas`.

Keep `tools/public_api_manifest.txt` in sync with the public `UI_*` API (`python3 tools/vba_api.py --write`). Code that changes VBA behaviour also needs `Test_EXCEL_UI_RunReleaseCertification` on a real Excel host; static checks cannot replace it.
